"""
parse.py  –  third draft

Parses a folder of court-watching .docx files into a flat CSV dataset.
One row per case; header metadata is shared across all cases in the file.

Usage:
    python parse.py                          # reads ./mini_data, writes output.csv
    python parse.py --input ./mini_data --output results.csv

Requires: pip install python-docx
"""

import re
import csv
import argparse
from pathlib import Path
from docx import Document
from collections import defaultdict

NaN = float("nan")

# ---------------------------------------------------------------------------
# Text normalisation
# ---------------------------------------------------------------------------

def clean(text: str) -> str:
    """Normalise whitespace and common unicode artifacts; never replace values."""
    return (text
            .replace("\u00a0", " ")
            .replace("\u2019", "'").replace("\u2018", "'")
            .replace("\u201c", '"').replace("\u201d", '"')
            .replace("\u2605", "").replace("\u2b51", "")
            .replace("✭", "").replace("⭑", "")
            .strip())


def para_text(para) -> str:
    return clean(para.text)


def nan_if_empty(val: str):
    """Return NaN for blank / whitespace-only strings, otherwise the string."""
    s = val.strip() if isinstance(val, str) else ""
    return s if s else NaN


# ---------------------------------------------------------------------------
# Document-level splitter: paragraphs → header block + per-case blocks
# ---------------------------------------------------------------------------

CASE_RE = re.compile(r"^\s*[*_]*Case\s*:\s*(\d+)", re.IGNORECASE)


def split_into_cases(doc: Document):
    """
    Returns:
        header_paras  : list of paragraphs before the first Case
        case_blocks   : list of (case_number_str, [paragraphs])
    """
    header_paras = []
    case_blocks  = []
    current_num  = None
    current_paras = []

    for para in doc.paragraphs:
        t = para_text(para)
        m = CASE_RE.match(t)
        if m:
            if current_num is not None:
                case_blocks.append((current_num, current_paras))
            current_num   = m.group(1)
            current_paras = [para]
        elif current_num is None:
            header_paras.append(para)
        else:
            current_paras.append(para)

    if current_num is not None:
        case_blocks.append((current_num, current_paras))

    return header_paras, case_blocks


# ---------------------------------------------------------------------------
# Section splitter within a case block (by roman numeral headings)
# ---------------------------------------------------------------------------

ROMAN_MAP = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6}
ROMAN_RE  = re.compile(r"^\s*(VI?|IV|I{1,3})\s*\.", re.IGNORECASE)


def split_into_sections(paragraphs):
    """Returns dict {0: [paras before §I], 1: [...], ..., 6: [...]}."""
    sections = {0: []}
    current  = 0

    for para in paragraphs:
        t = para_text(para)
        m = ROMAN_RE.match(t)
        if m:
            key = ROMAN_MAP.get(m.group(1).upper(), 0)
            current = key
            sections.setdefault(current, [])
        else:
            sections.setdefault(current, []).append(para)

    return sections


def section_lines(sections, num) -> list[str]:
    """Return list of non-empty paragraph texts for a section."""
    return [para_text(p) for p in sections.get(num, [])]


def section_blob(sections, num) -> str:
    return "\n".join(section_lines(sections, num))


# ---------------------------------------------------------------------------
# Field extraction helpers
# ---------------------------------------------------------------------------

def first_match(patterns, text, default=""):
    """
    Try each regex; return first capture group (stripped) or default.
    Uses MULTILINE so ^ / $ work per-line; no DOTALL so [^\n]* stays on one line.
    """
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE | re.MULTILINE)
        if m:
            return clean(m.group(1))
    return default


def collect_until(lines: list[str], start_idx: int, stop_re) -> str:
    """
    Starting at lines[start_idx+1], collect text until stop_re matches a line.
    Returns joined, collapsed-whitespace string.
    """
    parts = []
    for i in range(start_idx + 1, len(lines)):
        if stop_re and stop_re.search(lines[i]):
            break
        parts.append(lines[i])
    return re.sub(r"\s+", " ", " ".join(parts)).strip()


# ---------------------------------------------------------------------------
# Header parsing (paragraphs before Case: 1)
# ---------------------------------------------------------------------------

REPEATED_ARGS_RE = re.compile(
    r"write down any arguments repeated", re.IGNORECASE)
GENERAL_OBS_RE   = re.compile(
    r"general observations and takeaways", re.IGNORECASE)
COURTROOM_COND_RE = re.compile(r"courtroom conditions", re.IGNORECASE)
ABILITY_RE       = re.compile(r"ability to hear", re.IGNORECASE)
WINDOWS_RE       = re.compile(r"does the courtroom have windows", re.IGNORECASE)
COURTROOM_TYPE_RE = re.compile(
    r"what type of courtroom were the hearings held in", re.IGNORECASE)
NAV_RE           = re.compile(r"courthouse navigator", re.IGNORECASE)


def parse_header(paragraphs) -> dict:
    header = {
        "source_file"        : "",
        "documenter_name"    : "",
        "hearing_date"       : "",
        "arrived_at"         : "",
        "left_at"            : "",
        "judge_header"       : "",
        "courtroom_number"   : "",
        "repeated_arguments" : "",
        "general_observations": "",
        "courthouse_navigator": "",
        "ability_to_hear"    : "",
        "courtroom_has_windows": "",
        "courtroom_type"     : "",
    }

    lines = [para_text(p) for p in paragraphs]
    n = len(lines)

    # Sentinels: patterns that begin a new header field (used to stop multi-line captures)
    NEXT_HDR = re.compile(
        r"(does the courtroom have windows|what type of courtroom|"
        r"courthouse navigator|ability to hear|courtroom conditions|"
        r"general observations|write down any arguments|"
        r"arrived at court|left court|judge\s*:|courtroom\s*#|^date:)",
        re.IGNORECASE,
    )

    i = 0
    while i < n:
        t    = lines[i]
        tl   = t.lower()

        # ---- Simple single-line fields ----
        if tl.startswith("documenter name:"):
            header["documenter_name"] = t.split(":", 1)[1].strip()

        elif re.match(r"^date:", t, re.IGNORECASE) and not header["hearing_date"]:
            header["hearing_date"] = t.split(":", 1)[1].strip()

        elif "arrived at court house at" in tl:
            header["arrived_at"] = re.split(
                r"arrived at court house at\s*:", t, flags=re.IGNORECASE)[1].strip()

        elif "left court house at" in tl:
            header["left_at"] = re.split(
                r"left court house at\s*:", t, flags=re.IGNORECASE)[1].strip()

        elif re.match(r"^judge\s*:", t, re.IGNORECASE):
            header["judge_header"] = t.split(":", 1)[1].strip()

        elif re.match(r"^courtroom\s*#", t, re.IGNORECASE):
            m = re.search(r"courtroom\s*#\s*:?\s*(\S+)", t, re.IGNORECASE)
            if m:
                header["courtroom_number"] = m.group(1).rstrip()

        # ---- Courthouse Navigator (on a bullet line "- Name: X") ----
        elif NAV_RE.search(t):
            m = re.search(r"courthouse navigator\s*(?:name)?\s*:?\s*(.+)",
                           t, re.IGNORECASE)
            if m:
                header["courthouse_navigator"] = m.group(1).strip()

        # ---- Multi-line fields: collect bullets until next sentinel ----
        elif REPEATED_ARGS_RE.search(t):
            parts = []
            j = i + 1
            while j < n:
                if GENERAL_OBS_RE.search(lines[j]) or COURTROOM_COND_RE.search(lines[j]):
                    break
                stripped = lines[j].lstrip("-• \t").strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            header["repeated_arguments"] = " | ".join(parts)

        elif GENERAL_OBS_RE.search(t):
            parts = []
            j = i + 1
            while j < n:
                if COURTROOM_COND_RE.search(lines[j]) or ABILITY_RE.search(lines[j]):
                    break
                stripped = lines[j].lstrip("-• \t").strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            header["general_observations"] = " | ".join(parts)

        elif ABILITY_RE.search(t):
            inline = re.split(r"ability to hear\s*\??", t,
                               flags=re.IGNORECASE, maxsplit=1)[-1].strip().lstrip(":").strip()
            parts  = [inline] if inline else []
            j = i + 1
            while j < n:
                if NEXT_HDR.search(lines[j]) and not ABILITY_RE.search(lines[j]):
                    break
                stripped = lines[j].lstrip("-• \t").strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            header["ability_to_hear"] = " ".join(parts).strip()

        elif WINDOWS_RE.search(t):
            inline = re.split(r"does the courtroom have windows\s*\??", t,
                               flags=re.IGNORECASE, maxsplit=1)[-1].strip().lstrip(":").strip()
            parts  = [inline] if inline else []
            j = i + 1
            while j < n:
                if NEXT_HDR.search(lines[j]) and not WINDOWS_RE.search(lines[j]):
                    break
                stripped = lines[j].lstrip("-• \t").strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            header["courtroom_has_windows"] = " ".join(parts).strip()

        elif COURTROOM_TYPE_RE.search(t):
            inline = re.split(
                r"what type of courtroom were the hearings held in\s*\??",
                t, flags=re.IGNORECASE, maxsplit=1)[-1].strip().lstrip(":").strip()
            parts = [inline] if inline else []
            j = i + 1
            while j < n:
                if NEXT_HDR.search(lines[j]) and not COURTROOM_TYPE_RE.search(lines[j]):
                    break
                stripped = lines[j].lstrip("-• \t").strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            header["courtroom_type"] = " ".join(parts).strip()

        i += 1

    return header


# ---------------------------------------------------------------------------
# Case parsing – one section at a time
# ---------------------------------------------------------------------------

# Regexes for Section III boundary markers (stop conditions)
S3_WHERE_RE    = re.compile(r"^where\s*\?", re.IGNORECASE)
S3_CONTACT_RE  = re.compile(r"what reason did officers give for initiating contact", re.IGNORECASE)
S3_GUN_RE      = re.compile(r"was a gun found", re.IGNORECASE)
S3_GUN_LOC_RE  = re.compile(r"if yes.*where was the gun found|where was the gun found",
                              re.IGNORECASE)
S3_OTHER_RE    = re.compile(r"other important facts about the interaction with police",
                              re.IGNORECASE)

# Regexes for Section V boundary markers
S5_WHO_RE       = re.compile(r"who was there\?.*what was shared", re.IGNORECASE)
S5_JUDGE_CMT_RE = re.compile(r"what comments.*judge make about family", re.IGNORECASE)
S5_DEPENDANTS_RE = re.compile(
    r"does the accused person have children/spouse/other dependants/job",
    re.IGNORECASE)
S5_OTHER_RE     = re.compile(
    r"what else did you learn about the person", re.IGNORECASE)

# Narrative headers in Section IV
S4_STATE_RE   = re.compile(r"state'?s?\s+narrative", re.IGNORECASE)
S4_DEFENSE_RE = re.compile(r"defense'?s?\s+narrative", re.IGNORECASE)
S4_JUDGE_RE   = re.compile(r"judge'?s?\s+3\s+prongs?", re.IGNORECASE)

EDITOR_NOTE_RE = re.compile(r"\[.*?editor\'?s?\s+note", re.IGNORECASE)

def collect_narrative(paragraphs, header_re, stop_re=None):
    """
    Collect everything between narrative headers and reconstruct
    Word numbering hierarchy:

        ilvl=0 -> 1., 2., 3.
        ilvl=1 -> a., b., c.
        ilvl=2 -> i., ii., iii.

    Returns a single string.
    """

    collecting = False
    parts = []

    counters = defaultdict(int)

    for para in paragraphs:

        text = para_text(para)

        if not collecting:
            if header_re.search(text):
                collecting = True
            continue

        if stop_re and stop_re.search(text):
            break

        pPr = para._element.pPr

        if (
            pPr is not None
            and pPr.numPr is not None
            and pPr.numPr.ilvl is not None
        ):

            ilvl = int(pPr.numPr.ilvl.val)

            counters[ilvl] += 1

            # reset deeper levels
            for lvl in list(counters.keys()):
                if lvl > ilvl:
                    counters[lvl] = 0

            if ilvl == 0:
                prefix = f"{counters[0]}."

            elif ilvl == 1:
                prefix = f"{chr(96 + counters[1])}."

            elif ilvl == 2:
                roman = [
                    "i", "ii", "iii", "iv", "v",
                    "vi", "vii", "viii", "ix", "x"
                ]
                idx = counters[2] - 1
                prefix = (
                    f"{roman[idx]}."
                    if 0 <= idx < len(roman)
                    else f"{counters[2]}."
                )

            else:
                prefix = "-"

            parts.append(f"{prefix} {text}")

        else:
            parts.append(text)

    return " ".join(parts).strip()


def parse_case(case_num: str, paragraphs: list) -> dict:
    secs  = split_into_sections(paragraphs)

    # Collect preamble lines between "Case: N" and Section I, excluding the heading itself.
    preamble_lines = [para_text(p) for p in secs.get(0, [])]
    preamble_lines = [l for l in preamble_lines if not CASE_RE.match(l)]
    # Any Editor's Notes in the preamble are appended to time_hearing_began below.
    preamble_editor_notes = " ".join(
        l for l in preamble_lines if EDITOR_NOTE_RE.search(l)
    ).strip()

    # ------------------------------------------------------------------ §I
    s1_lines = section_lines(secs, 1)
    s1 = "\n".join(s1_lines)

    time_began_base = first_match([
        r"time hearing began\s*:?\s*(.+?)(?:\s{2,}|\s*time ended)",
        r"time hearing began\s*:[^\S\n]*(\S[^\n]*)$",
    ], s1)
    time_began = (time_began_base + (" " + preamble_editor_notes if preamble_editor_notes else "")).strip()
    time_ended = first_match([r"time ended\s*:[^\S\n]*(\S[^\n]*)$"], s1)

    judge_case         = first_match([r"^judge\s*:[^\S\n]*(\S[^\n]*)$"], s1)
    prosecutor         = first_match([r"^prosecutor\s*:[^\S\n]*(\S[^\n]*)$"], s1)
    defense_attorney   = first_match([r"^defense attorney\s*:[^\S\n]*(\S[^\n]*)$"], s1)

    # Defense attorney type: strip the long question text if present
    # Variants seen:
    #   "Is the defense attorney a Public Defender, private attorney, or other: Unknown"
    #   "Defense attorney is a: Public Defender"
    #   "Defense attorney is a Public Defender."   (no colon)
    #   "The defense attorney is a: Public Defender"
    #   "The defense attorney is a Public Defender"
    def_type_raw = first_match([
        # "Is the defense attorney a ... : <answer>"
        r"is the defense attorney a[^:]*:\s*\*?\*?(\S[^\n]*)$",
        # "Defense attorney is a: <answer>"  or  "The defense attorney is a: <answer>"
        r"(?:the\s+)?defense attorney is an?\s*:\s*\*?\*?(\S[^\n]*)$",
        # "Defense attorney is a Public Defender."  (no colon before answer)
        r"(?:the\s+)?defense attorney is an?\s+(\S[^\n]*)$",
    ], s1)
    # Strip any residual "Public Defender, private attorney, or other" preamble
    def_type_raw = re.sub(
        r"^public defender,\s*private attorney,\s*or other\s*:\s*",
        "", def_type_raw, flags=re.IGNORECASE,
    ).strip(" *.")
    defense_attorney_type = def_type_raw

    accused_initials = first_match([
        r"accused person(?:\s+full\s+name)?\s*:\s*(\S[^\n]*)$",
    ], s1)
    gender_presentation = first_match([
        r"gender presentation\s*:\s*\*?\*?(\S[^\n]*?)\*?\*?$",
    ], s1)

    # Perceived race: capture the matching line, then append any immediately
    # following [Editor's Note: ...] paragraph (scan past blank lines)
    perceived_race_base = first_match([
        r"perceived race\s*:\s*\*?\*?(\S[^\n]*?)\*?\*?$",
    ], s1)
    editor_note_for_race = ""
    found_race = False
    for line in s1_lines:
        if found_race:
            stripped = line.strip()
            if not stripped:
                continue   # skip blank lines between perceived race and editor note
            if re.match(r"\[.*editor'?s? note", stripped, re.IGNORECASE):
                editor_note_for_race = " " + stripped
            break
        if re.search(r"perceived race", line, re.IGNORECASE):
            found_race = True
    perceived_race = (perceived_race_base + editor_note_for_race).strip()

    pretrial_fta   = first_match([r"failure to appear\s*:\s*(\S[^\n]*)$"], s1)
    pretrial_nca   = first_match([r"new criminal activity\s*:\s*(\S[^\n]*)$"], s1)
    pretrial_score = first_match([
        r"(?:pre-?trial\s+)?supervision score\s*(?:level\s*)?rec(?:ommendation)?\s*:\s*(\S[^\n]*)$",
    ], s1)

    # ------------------------------------------------------------------ §II
    s2 = section_blob(secs, 2)
    primary_charge     = first_match([r"primary charge\s*:\s*\*?\*?\s*(\S[^\n]*)$"], s2)
    additional_charges = first_match([r"any additional charges\s*:\s*\*?\*?\s*(\S[^\n]*)$"], s2)

    # ------------------------------------------------------------------ §III
    s3_lines = section_lines(secs, 3)

    # arrest_datetime: value after "When did this occur...?" on the SAME line only.
    # If nothing follows the question mark, leave blank.
    arrest_datetime = ""
    arrest_location = ""
    reason_for_contact = ""
    gun_found = ""
    gun_location = ""
    other_arrest_facts = ""

    i = 0
    while i < len(s3_lines):
        line = s3_lines[i]

        if re.search(r"when did this occur", line, re.IGNORECASE):
            # value is everything after the "?"
            m = re.search(r"when did this occur[^?]*\?\s*(\S.*)?$",
                           line, re.IGNORECASE)
            if m and m.group(1):
                arrest_datetime = m.group(1).strip()
            # else blank – don't grab next line

        elif S3_WHERE_RE.match(line):
            # value after "Where? " on the SAME line only
            m = re.search(r"where\s*\?\s*(\S.*)?$", line, re.IGNORECASE)
            if m and m.group(1):
                arrest_location = m.group(1).strip()
            # else blank – do NOT grab next line (which is the contact question)

        elif S3_CONTACT_RE.search(line):
            # Inline value on same line (e.g. Case 3, 7, 8)?
            m = re.search(
                r"what reason did officers give for initiating contact\?\s*(\S.+)?$",
                line, re.IGNORECASE)
            inline = (m.group(1) or "").strip() if m else ""
            parts = [inline] if inline else []
            j = i + 1
            while j < len(s3_lines):
                nxt = s3_lines[j]
                if S3_GUN_RE.search(nxt):
                    break
                stripped = nxt.lstrip("-• \t").strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            reason_for_contact = re.sub(r"\s+", " ", " ".join(parts)).strip()

        elif S3_GUN_RE.search(line):
            m = re.search(r"was a gun found[^:]*:\s*\*?\*?(\S[^\n]*)$",
                           line, re.IGNORECASE)
            if m:
                gun_found = m.group(1).strip().rstrip("* ")
            else:
                # answer might be on the very next line
                if i + 1 < len(s3_lines) and not S3_GUN_LOC_RE.search(s3_lines[i+1]):
                    candidate = s3_lines[i+1].lstrip("-• \t").strip()
                    if candidate and not re.search(
                            r"if yes|where was", candidate, re.IGNORECASE):
                        gun_found = candidate

        elif S3_GUN_LOC_RE.search(line):
            # collect the immediately following non-empty line(s) until next sentinel
            parts = []
            j = i + 1
            while j < len(s3_lines):
                nxt = s3_lines[j]
                if S3_OTHER_RE.search(nxt) or S3_GUN_RE.search(nxt):
                    break
                stripped = nxt.lstrip("-• \t").strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            gun_location = re.sub(r"\s+", " ", " ".join(parts)).strip()

        elif S3_OTHER_RE.search(line):
            parts = []
            j = i + 1
            while j < len(s3_lines):
                stripped = s3_lines[j].lstrip("-• \t").strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            other_arrest_facts = re.sub(r"\s+", " ", " ".join(parts)).strip()

        i += 1

    # ------------------------------------------------------------------ §IV
    # Use ALL paragraph texts including blank lines so collect_narrative
    # captures everything between narrative headers without gaps.

    # ------------------------------------------------------------------ §IV

    s4_paras = secs.get(4, [])

    state_narrative = collect_narrative(
        s4_paras,
        S4_STATE_RE,
        S4_DEFENSE_RE
    )

    defense_narrative = collect_narrative(
        s4_paras,
        S4_DEFENSE_RE,
        S4_JUDGE_RE
    )

    judge_3_prongs = collect_narrative(
        s4_paras,
        S4_JUDGE_RE,
        None
    )

    # ------------------------------------------------------------------ §V
    s5_lines = section_lines(secs, 5)

    family_present          = ""
    family_who              = ""
    judge_comments_on_family= ""
    dependants_spouse_job              = ""
    other_info_about_person = ""

    i = 0
    while i < len(s5_lines):
        line = s5_lines[i]

        if re.search(r"were family/friends present and acknowledged", line, re.IGNORECASE):
            # value may be inline after "?" — require a real answer (not just punctuation)
            m = re.search(
                r"were family/friends present and acknowledged\?[^\S\n]*(\S[^\n]*)$",
                line, re.IGNORECASE)
            inline_val = ""
            if m:
                candidate = m.group(1).strip().rstrip("*")
                # Reject bare punctuation artefacts from the label itself
                if candidate and candidate not in ("?", "*", "**"):
                    inline_val = candidate
            if inline_val:
                family_present = inline_val
            else:
                # Look at the next non-blank line
                j = i + 1
                while j < len(s5_lines):
                    candidate = s5_lines[j].strip()
                    if candidate:
                        if not S5_WHO_RE.search(candidate):
                            family_present = candidate
                        break
                    j += 1

        elif S5_WHO_RE.search(line):
            parts = []
            j = i + 1
            while j < len(s5_lines):
                if S5_JUDGE_CMT_RE.search(s5_lines[j]):
                    break
                stripped = s5_lines[j].strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            family_who = re.sub(r"\s+", " ", " ".join(parts)).strip()

        elif S5_JUDGE_CMT_RE.search(line):
            parts = []
            j = i + 1
            while j < len(s5_lines):
                if S5_DEPENDANTS_RE.search(s5_lines[j]):
                    break
                stripped = s5_lines[j].strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            judge_comments_on_family = re.sub(r"\s+", " ", " ".join(parts)).strip()

        elif S5_DEPENDANTS_RE.search(line):

            PROMPT = "does the accused person have children/spouse/other dependants/job"

            collected = []

            # capture anything appearing after the question on the same line
            m = re.search(
                r"does the accused person have children/spouse/other dependants/job\??\s*(.*)$",
                line,
                re.IGNORECASE,
            )

            if m:
                remainder = m.group(1).strip(" :*")
                if remainder:
                    collected.append(remainder)

            # capture subsequent lines until the next Section V question
            j = i + 1
            while j < len(s5_lines):

                nxt = s5_lines[j].strip()

                if not nxt:
                    j += 1
                    continue

                if S5_OTHER_RE.search(nxt):
                    break

                collected.append(nxt)
                j += 1

            dependants_spouse_job = re.sub(
                r"\s+",
                " ",
                " ".join(collected)
            ).strip()

        elif S5_OTHER_RE.search(line):
            parts = []
            j = i + 1
            while j < len(s5_lines):
                stripped = s5_lines[j].strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            other_info_about_person = re.sub(r"\s+", " ", " ".join(parts)).strip()

        i += 1

    # ------------------------------------------------------------------ §VI
    s6_lines = section_lines(secs, 6)
    # Remove trailing "Note to reader" boilerplate
    outcome_parts = []
    for line in s6_lines:
        if re.match(r"note to reader", line, re.IGNORECASE):
            break
        outcome_parts.append(line)
    outcome = re.sub(r"\s+", " ", " ".join(outcome_parts)).strip()

    # ------------------------------------------------------------------ assemble
    return {
        "case_number"             : case_num,
        # §I
        "time_hearing_began"      : time_began,
        "time_hearing_ended"      : time_ended,
        "judge_case"              : judge_case,
        "prosecutor"              : prosecutor,
        "defense_attorney"        : defense_attorney,
        "defense_attorney_type"   : defense_attorney_type,
        "accused_person_initials" : accused_initials,
        "accused_gender_presentation": gender_presentation,
        "accused_perceived_race"  : perceived_race,
        "pretrial_failure_to_appear": pretrial_fta,
        "pretrial_new_criminal_activity": pretrial_nca,
        "pretrial_supervision_score": pretrial_score,
        # §II
        "primary_charge"          : primary_charge,
        "additional_charges"      : additional_charges,
        # §III
        "arrest_datetime"         : arrest_datetime,
        "arrest_location"         : arrest_location,
        "reason_for_contact"      : reason_for_contact,
        "gun_found"               : gun_found,
        "gun_location"            : gun_location,
        "other_arrest_facts"      : other_arrest_facts,
        # §IV
        "state_narrative"         : state_narrative,
        "defense_narrative"       : defense_narrative,
        "judge_3_prongs"          : judge_3_prongs,
        # §V
        "family_present"          : family_present,
        "family_who"              : family_who,
        "judge_comments_on_family": judge_comments_on_family,
        "dependants_spouse_job"   : dependants_spouse_job,
        "other_info_about_person" : other_info_about_person,
        # §VI
        "outcome"                 : outcome,
    }


# ---------------------------------------------------------------------------
# File processor
# ---------------------------------------------------------------------------

FIELDNAMES = [
    # Session metadata
    "source_file", "documenter_name", "hearing_date",
    "arrived_at", "left_at",
    "judge_header", "courtroom_number",
    "repeated_arguments", "general_observations",
    "courthouse_navigator", "ability_to_hear",
    "courtroom_has_windows", "courtroom_type",
    # Case
    "case_number",
    "time_hearing_began", "time_hearing_ended",
    "judge_case", "prosecutor", "defense_attorney", "defense_attorney_type",
    "accused_person_initials", "accused_gender_presentation", "accused_perceived_race",
    "pretrial_failure_to_appear", "pretrial_new_criminal_activity",
    "pretrial_supervision_score",
    "primary_charge", "additional_charges",
    "arrest_datetime", "arrest_location", "reason_for_contact",
    "gun_found", "gun_location", "other_arrest_facts",
    "state_narrative", "defense_narrative", "judge_3_prongs",
    "family_present", "family_who", "judge_comments_on_family",
    "dependants_spouse_job", "other_info_about_person",
    "outcome",
]


def process_file(path: Path) -> list[dict]:
    doc = Document(path)
    header_paras, case_blocks = split_into_cases(doc)
    header = parse_header(header_paras)
    header["source_file"] = path.name

    rows = []
    for case_num, paras in case_blocks:
        case_data = parse_case(case_num, paras)
        row = {
            "source_file"          : header["source_file"],
            "documenter_name"      : header["documenter_name"],
            "hearing_date"         : header["hearing_date"],
            "arrived_at"           : header["arrived_at"],
            "left_at"              : header["left_at"],
            "judge_header"         : header["judge_header"],
            "courtroom_number"     : header["courtroom_number"],
            "repeated_arguments"   : header["repeated_arguments"],
            "general_observations" : header["general_observations"],
            "courthouse_navigator" : header["courthouse_navigator"],
            "ability_to_hear"      : header["ability_to_hear"],
            "courtroom_has_windows": header["courtroom_has_windows"],
            "courtroom_type"       : header["courtroom_type"],
            **case_data,
        }
        # Apply NaN to all blank string values
        for k, v in row.items():
            if isinstance(v, str):
                row[k] = nan_if_empty(v)
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Parse court watching docx files to CSV.")
    parser.add_argument("--input",  default="mini_data")
    parser.add_argument("--output", default="output.csv")
    args = parser.parse_args()

    input_dir  = Path(args.input)
    docx_files = sorted(input_dir.glob("*.docx"))

    if not docx_files:
        print(f"No .docx files found in '{input_dir}'. Exiting.")
        return

    all_rows = []
    for f in docx_files:
        print(f"  Processing: {f.name}")
        try:
            rows = process_file(f)
            print(f"    → {len(rows)} case(s) found")
            all_rows.extend(rows)
        except Exception as e:
            import traceback
            print(f"    ERROR: {e}")
            traceback.print_exc()

    output_path = Path(args.output)
    with open(output_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=FIELDNAMES, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(all_rows)

    print(f"\nDone. {len(all_rows)} total rows written to '{output_path}'.")


if __name__ == "__main__":
    main()