"""
parse.py

Parses a folder of court watching .docx files into a flat CSV dataset.
Each row represents one case. Header-level metadata is shared across all
cases in the same document.

Usage:
    python parse.py                        # reads ./mini_data, writes output.csv
    python parse.py --input ./mini_data --output results.csv

Requires: pip install python-docx
"""

import re
import csv
import argparse
from pathlib import Path
from docx import Document


# ---------------------------------------------------------------------------
# Text cleaning
# ---------------------------------------------------------------------------

def clean(text: str) -> str:
    """Normalise whitespace and common unicode artifacts; do NOT replace values."""
    return (text
            .replace("\u00a0", " ")   # non-breaking space
            .replace("\u2019", "'").replace("\u2018", "'")
            .replace("\u201c", '"').replace("\u201d", '"')
            .replace("\u2605", "").replace("\u2b51", "")  # star decorators
            .replace("✭", "")
            .strip())


def para_text(para) -> str:
    return clean(para.text)


# ---------------------------------------------------------------------------
# Document splitter: paragraphs → header block + per-case blocks
# ---------------------------------------------------------------------------

CASE_RE = re.compile(r"^\s*[*_]?Case\s*:\s*\d+", re.IGNORECASE)

def split_into_cases(doc: Document):
    """
    Returns:
        header_paras  : list of paragraphs before the first Case
        case_blocks   : list of (case_number_str, [paragraphs])
    """
    header_paras = []
    case_blocks = []
    current_num = None
    current_paras = []

    for para in doc.paragraphs:
        t = para_text(para)
        m = CASE_RE.match(t)
        if m:
            if current_num is not None:
                case_blocks.append((current_num, current_paras))
            num_m = re.search(r"\d+", t)
            current_num = num_m.group(0) if num_m else ""
            current_paras = [para]
        elif current_num is None:
            header_paras.append(para)
        else:
            current_paras.append(para)

    if current_num is not None:
        case_blocks.append((current_num, current_paras))

    return header_paras, case_blocks


# ---------------------------------------------------------------------------
# Section splitter within a case block
# ---------------------------------------------------------------------------

# Matches "I.", "II.", "III.", "IV.", "V.", "VI." at the start of a paragraph
ROMAN_RE = re.compile(r"^\s*(I{1,3}|IV|VI?)\s*\.", re.IGNORECASE)

def split_into_sections(paragraphs):
    """
    Split case paragraphs into a dict keyed by roman numeral index (1-6).
    Section 0 = preamble before Section I.
    """
    sections = {0: []}
    current = 0
    roman_map = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6}

    for para in paragraphs:
        t = para_text(para)
        m = ROMAN_RE.match(t)
        if m:
            key = roman_map.get(m.group(1).upper(), 0)
            current = key
            sections.setdefault(current, [])
        else:
            sections.setdefault(current, []).append(para)

    return sections


def section_text(sections, num):
    """Return all paragraph texts for a section joined by newline."""
    return "\n".join(para_text(p) for p in sections.get(num, []))


# ---------------------------------------------------------------------------
# Header parsing (before Case: 1)
# ---------------------------------------------------------------------------

def parse_header(paragraphs) -> dict:
    """
    Walk header paragraphs sequentially, capturing fields.
    Multi-line fields (ability_to_hear, courtroom_type) consume lines
    until the next recognised field header.
    """
    header = {
        "source_file": "",
        "documenter_name": "",
        "hearing_date": "",
        "arrived_at": "",
        "left_at": "",
        "judge_header": "",
        "courtroom_number": "",
        "courthouse_navigator": "",
        "ability_to_hear": "",
        "courtroom_has_windows": "",
        "courtroom_type": "",
    }

    lines = [para_text(p) for p in paragraphs]

    # Sentinel patterns that mark the START of the next field
    NEXT_FIELD_RE = re.compile(
        r"(does the courtroom have windows|what type of courtroom|"
        r"arrived at court|left court|judge\s*:|courtroom\s*#|date:|"
        r"courthouse navigator|ability to hear)",
        re.IGNORECASE,
    )

    i = 0
    while i < len(lines):
        t = lines[i]
        tl = t.lower()

        if tl.startswith("documenter name:"):
            header["documenter_name"] = t.split(":", 1)[1].strip()

        elif re.match(r"^date:", t, re.IGNORECASE) and not header["hearing_date"]:
            header["hearing_date"] = t.split(":", 1)[1].strip()

        elif "arrived at court house at" in tl:
            header["arrived_at"] = re.split(r"arrived at court house at\s*:", t, flags=re.IGNORECASE)[1].strip()

        elif "left court house at" in tl:
            header["left_at"] = re.split(r"left court house at\s*:", t, flags=re.IGNORECASE)[1].strip()

        elif re.match(r"^judge\s*:", t, re.IGNORECASE):
            header["judge_header"] = t.split(":", 1)[1].strip()

        elif re.match(r"^courtroom\s*#", t, re.IGNORECASE):
            m = re.search(r"courtroom\s*#\s*:?\s*(\S+)", t, re.IGNORECASE)
            if m:
                header["courtroom_number"] = m.group(1)

        elif "courthouse navigator" in tl:
            m = re.search(r"courthouse navigator\s*(?:name)?\s*:?\s*(.+)", t, re.IGNORECASE)
            if m:
                header["courthouse_navigator"] = m.group(1).strip()

        elif "ability to hear" in tl:
            # Collect the inline value (after the ?) plus any following bullet lines
            # until we hit the next recognised field
            inline = re.split(r"ability to hear\s*\??", t, flags=re.IGNORECASE, maxsplit=1)[-1].strip().lstrip(":").strip()
            parts = [inline] if inline else []
            j = i + 1
            while j < len(lines):
                nxt = lines[j]
                if NEXT_FIELD_RE.search(nxt):
                    break
                stripped = nxt.lstrip("-• \t").strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            header["ability_to_hear"] = " ".join(parts).strip()

        elif "does the courtroom have windows" in tl:
            # Same pattern: inline + following bullet lines until next field
            inline = re.split(r"does the courtroom have windows\s*\??", t, flags=re.IGNORECASE, maxsplit=1)[-1].strip().lstrip(":").strip()
            parts = [inline] if inline else []
            j = i + 1
            while j < len(lines):
                nxt = lines[j]
                if NEXT_FIELD_RE.search(nxt):
                    break
                stripped = nxt.lstrip("-• \t").strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            header["courtroom_has_windows"] = " ".join(parts).strip()

        elif "what type of courtroom" in tl:
            inline = re.split(r"what type of courtroom were the hearings held in\s*\??", t, flags=re.IGNORECASE, maxsplit=1)[-1].strip().lstrip(":").strip()
            parts = [inline] if inline else []
            j = i + 1
            while j < len(lines):
                nxt = lines[j]
                if NEXT_FIELD_RE.search(nxt):
                    break
                stripped = nxt.lstrip("-• \t").strip()
                if stripped:
                    parts.append(stripped)
                j += 1
            header["courtroom_type"] = " ".join(parts).strip()

        i += 1

    return header


# ---------------------------------------------------------------------------
# Case parsing — one section at a time
# ---------------------------------------------------------------------------

def first_match(patterns, text, default="", multiline_ok=False):
    """Try each pattern; return first capture group or default.
    By default patterns are single-line (no DOTALL). Pass multiline_ok=True
    for patterns that intentionally span lines.
    """
    flags = re.IGNORECASE | re.MULTILINE
    if multiline_ok:
        flags |= re.DOTALL
    for pat in patterns:
        m = re.search(pat, text, flags)
        if m:
            return clean(m.group(1))
    return default


def parse_case(case_num: str, paragraphs: list) -> dict:
    secs = split_into_sections(paragraphs)

    # ---- Section I: People & Case Information ----
    s1 = section_text(secs, 1)

    time_began = first_match([
        r"time hearing began\s*:?\s*(.+?)(?:\s{2,}|\s*time ended)",
        r"time hearing began\s*:?\s*([^\n]+)",
    ], s1)
    time_ended = first_match([
        r"time ended\s*:?\s*([^\n]+)",
    ], s1)
    judge_case = first_match([r"^judge\s*:[^\S\n]*(\S[^\n]*)$"], s1)
    prosecutor = first_match([r"^prosecutor\s*:[^\S\n]*(\S[^\n]*)$"], s1)
    defense_attorney = first_match([r"^defense attorney\s*:[^\S\n]*(\S[^\n]*)$"], s1)
    defense_attorney_type = first_match([
        r"(?:defense attorney is a|is the defense attorney a|the defense attorney is an?)\s*:?\s*\*?\*?([^\n]+?)\*?\*?$",
    ], s1)
    accused_initials = first_match([
        r"accused person(?:\s+full\s+name)?\s*:\s*([^\n]+)",
    ], s1)
    gender_presentation = first_match([r"gender presentation\s*:\s*\*?\*?([^\n]+?)\*?\*?$"], s1)
    perceived_race = first_match([r"perceived race\s*:\s*\*?\*?([^\n]+?)\*?\*?$"], s1)
    pretrial_fta = first_match([r"failure to appear\s*:\s*([^\n]+)"], s1)
    pretrial_nca = first_match([r"new criminal activity\s*:\s*([^\n]+)"], s1)
    pretrial_score = first_match([
        r"(?:pre-?trial\s+)?supervision score\s*(?:level\s*)?rec(?:ommendation)?\s*:\s*([^\n]+)",
    ], s1)

    # ---- Section II: Charges ----
    s2 = section_text(secs, 2)

    primary_charge = first_match([
        r"primary charge\s*:\s*\*?\*?\s*([^\n]+)",
    ], s2)
    additional_charges = first_match([
        r"any additional charges\s*:\s*\*?\*?\s*([^\n]+)",
    ], s2)

    # ---- Section III: Facts of the Arrest ----
    s3 = section_text(secs, 3)

    arrest_datetime = first_match([
        r"when did this occur[^?]*\?\s*([^\n]+)",
    ], s3)
    arrest_location = first_match([
        r"where\s*\?\s*([^\n]+)",
    ], s3)
    # Reason: everything from the label until the next bullet label
    contact_m = re.search(
        r"what reason did officers give for initiating contact\??\s*\n(.*?)(?=\n\s*-?\s*was a gun found|\Z)",
        s3, re.IGNORECASE | re.DOTALL,
    )
    reason_for_contact = re.sub(r"\s+", " ", contact_m.group(1)).strip() if contact_m else ""

    gun_found = first_match([
        r"was a gun found[^:]*:\s*\*?\*?([^\n]+?)\*?\*?$",
        r"was a gun found[^\n]*\n\s*\*?\*?([^\n]+?)\*?\*?$",
    ], s3)
    gun_location = first_match([
        r"(?:if yes,?\s*)?where was the gun found\??\s*\n\s*[-•]?\s*\*?\*?([^\n]+?)\*?\*?$",
    ], s3)
    other_facts_m = re.search(
        r"other important facts about the interaction with police\??\s*\n(.*?)(?=\Z)",
        s3, re.IGNORECASE | re.DOTALL,
    )
    other_arrest_facts = re.sub(r"\s+", " ", other_facts_m.group(1)).strip() if other_facts_m else ""

    # ---- Section IV: Three Prongs + Narratives ----
    # The boilerplate "1: ... 2: ... 3: ..." lines are skipped.
    # We capture State's Narrative, Defense's Narrative, Judge's 3 Prongs.
    s4 = section_text(secs, 4)

    state_narrative_m = re.search(
        r"state'?s? narrative\s*:?\s*\n(.*?)(?=\n\s*defense'?s? narrative|\Z)",
        s4, re.IGNORECASE | re.DOTALL,
    )
    state_narrative = re.sub(r"\s+", " ", state_narrative_m.group(1)).strip() if state_narrative_m else ""

    defense_narrative_m = re.search(
        r"defense'?s? narrative\s*:?\s*\n(.*?)(?=\n\s*judge'?s? 3 prongs|\Z)",
        s4, re.IGNORECASE | re.DOTALL,
    )
    defense_narrative = re.sub(r"\s+", " ", defense_narrative_m.group(1)).strip() if defense_narrative_m else ""

    judge_prongs_m = re.search(
        r"judge'?s? 3 prongs\s*:?\s*\n(.*?)(?=\Z)",
        s4, re.IGNORECASE | re.DOTALL,
    )
    judge_3_prongs = re.sub(r"\s+", " ", judge_prongs_m.group(1)).strip() if judge_prongs_m else ""

    # ---- Section V: Family & Community Presence ----
    s5 = section_text(secs, 5)

    family_present = first_match([
        r"were family/friends present and acknowledged\??\s*\*?\*?([^\n]+?)\*?\*?$",
    ], s5)

    family_who_m = re.search(
        r"who was there\? what was shared about the people present\?\s*\n(.*?)(?=\n\s*what comments|\Z)",
        s5, re.IGNORECASE | re.DOTALL,
    )
    family_who = re.sub(r"\s+", " ", family_who_m.group(1)).strip() if family_who_m else ""

    judge_family_m = re.search(
        r"what comments[^?]*judge make about family[^?]*\?\s*\n(.*?)(?=\n\s*does the accused|\Z)",
        s5, re.IGNORECASE | re.DOTALL,
    )
    judge_comments_on_family = re.sub(r"\s+", " ", judge_family_m.group(1)).strip() if judge_family_m else ""

    dependants = first_match([
        r"does the accused person have children[^:]*:\s*\*?\*?([^\n]+?)\*?\*?$",
    ], s5)

    other_info_m = re.search(
        r"what else did you learn about the person[^:]*:\s*\*?\*?\s*\n(.*?)(?=\Z)",
        s5, re.IGNORECASE | re.DOTALL,
    )
    other_info_about_person = re.sub(r"\s+", " ", other_info_m.group(1)).strip() if other_info_m else ""

    # ---- Section VI: Outcome ----
    s6 = section_text(secs, 6)
    # Take everything; trim trailing boilerplate "Note to reader" footer
    outcome = re.sub(r"\s+", " ", s6).strip()
    outcome = re.sub(
        r"\s*Note to reader\s*:.*$", "", outcome, flags=re.IGNORECASE | re.DOTALL
    ).strip()

    return {
        # Section I
        "case_number": case_num,
        "time_hearing_began": time_began,
        "time_hearing_ended": time_ended,
        "judge_case": judge_case,
        "prosecutor": prosecutor,
        "defense_attorney": defense_attorney,
        "defense_attorney_type": defense_attorney_type,
        "accused_person_initials": accused_initials,
        "accused_gender_presentation": gender_presentation,
        "accused_perceived_race": perceived_race,
        "pretrial_failure_to_appear": pretrial_fta,
        "pretrial_new_criminal_activity": pretrial_nca,
        "pretrial_supervision_score": pretrial_score,
        # Section II
        "primary_charge": primary_charge,
        "additional_charges": additional_charges,
        # Section III
        "arrest_datetime": arrest_datetime,
        "arrest_location": arrest_location,
        "reason_for_contact": reason_for_contact,
        "gun_found": gun_found,
        "gun_location": gun_location,
        "other_arrest_facts": other_arrest_facts,
        # Section IV
        "state_narrative": state_narrative,
        "defense_narrative": defense_narrative,
        "judge_3_prongs": judge_3_prongs,
        # Section V
        "family_present": family_present,
        "family_who": family_who,
        "judge_comments_on_family": judge_comments_on_family,
        "dependants": dependants,
        "other_info_about_person": other_info_about_person,
        # Section VI
        "outcome": outcome,
    }


# ---------------------------------------------------------------------------
# File processor
# ---------------------------------------------------------------------------

def process_file(path: Path) -> list[dict]:
    doc = Document(path)
    header_paras, case_blocks = split_into_cases(doc)
    header = parse_header(header_paras)
    header["source_file"] = path.name

    rows = []
    for case_num, paras in case_blocks:
        case_data = parse_case(case_num, paras)
        row = {
            # Header fields first, in document order
            "source_file": header["source_file"],
            "documenter_name": header["documenter_name"],
            "hearing_date": header["hearing_date"],
            "arrived_at": header["arrived_at"],
            "left_at": header["left_at"],
            "judge_header": header["judge_header"],
            "courtroom_number": header["courtroom_number"],
            "courthouse_navigator": header["courthouse_navigator"],
            "ability_to_hear": header["ability_to_hear"],
            "courtroom_has_windows": header["courtroom_has_windows"],
            "courtroom_type": header["courtroom_type"],
            **case_data,
        }
        rows.append(row)
    return rows


FIELDNAMES = [
    # Session metadata (header)
    "source_file", "documenter_name", "hearing_date",
    "arrived_at", "left_at",
    "judge_header", "courtroom_number", "courthouse_navigator",
    "ability_to_hear", "courtroom_has_windows", "courtroom_type",
    # Section I
    "case_number",
    "time_hearing_began", "time_hearing_ended",
    "judge_case", "prosecutor", "defense_attorney", "defense_attorney_type",
    "accused_person_initials", "accused_gender_presentation", "accused_perceived_race",
    "pretrial_failure_to_appear", "pretrial_new_criminal_activity", "pretrial_supervision_score",
    # Section II
    "primary_charge", "additional_charges",
    # Section III
    "arrest_datetime", "arrest_location", "reason_for_contact",
    "gun_found", "gun_location", "other_arrest_facts",
    # Section IV
    "state_narrative", "defense_narrative", "judge_3_prongs",
    # Section V
    "family_present", "family_who", "judge_comments_on_family",
    "dependants", "other_info_about_person",
    # Section VI
    "outcome",
]


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Parse court watching docx files to CSV.")
    parser.add_argument("--input", default="mini_data",
                        help="Folder containing .docx files (default: ./mini_data)")
    parser.add_argument("--output", default="output.csv",
                        help="Output CSV path (default: output.csv)")
    args = parser.parse_args()

    input_dir = Path(args.input)
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