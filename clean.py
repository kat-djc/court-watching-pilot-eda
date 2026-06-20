import re
import numpy as np
import pandas as pd

df = pd.read_csv("output.csv")

# ============================================================
# Helpers
# ============================================================

NULL_VALUES = {
    "unknown",
    "unsure",
    "na",
    "n/a",
    "nan",
    "[inaudible]",
    "public defender",
    "unknown, public defender",
    "unknown public defender",
    "unknown",
    "n/a",
    "na",
    "[not stated]",
    "not stated",
    "",
}

def normalize_text(val):

    if pd.isna(val):
        return None

    text = str(val).lower().strip()

    # collapse whitespace
    text = re.sub(r"\s+", " ", text)

    # remove trailing punctuation
    text = text.rstrip(".:,;!?")

    return text


def is_null_value(text):
    return text is None or text in NULL_VALUES


def match_patterns(val, patterns):

    text = normalize_text(val)

    if is_null_value(text):
        return []

    matches = []

    for clean_value, variants in patterns.items():
        if any(v in text for v in variants):
            matches.append(clean_value)

    return list(dict.fromkeys(matches))


# ============================================================
# Source file
# ============================================================

SOURCE_FILE_PATTERNS = {
    "Overflow": ["overflow"],
    "Room 102": ["room 102"],
}


def clean_source_file(val):
    return match_patterns(val, SOURCE_FILE_PATTERNS)


# ============================================================
# Time extraction
# ============================================================

TIME_RE = re.compile(
    r"(\d{1,2}:\d{2}\s*(?:a\.?m\.?|p\.?m\.?|am|pm)?|\d{1,2}\s*(?:a\.?m\.?|p\.?m\.?|am|pm))",
    re.IGNORECASE,
)


def extract_datetime_components(val):

    text = normalize_text(val)

    if is_null_value(text):
        return pd.NaT, pd.NaT

    # ------------------------
    # 1. Extract TIME
    # ------------------------
    time_match = TIME_RE.search(text)

    if time_match:
        time_text = time_match.group(1).lower()

        # normalize am/pm formatting
        time_text = (
            time_text
            .replace("a.m.", "am")
            .replace("p.m.", "pm")
            .replace("a.m", "am")
            .replace("p.m", "pm")
            .replace(" ", "")
        )

        # handle "9pm" → "9:00pm"
        time_text = re.sub(r"^(\d{1,2})(am|pm)$", r"\1:00\2", time_text)

        try:
            time_clean = pd.to_datetime(time_text, format="%I:%M%p").time()
        except Exception:
            time_clean = pd.NaT
    else:
        time_clean = pd.NaT

    # ------------------------
    # 2. Extract DATE
    # ------------------------
    try:
        date_clean = date_parser.parse(text, fuzzy=True).date()
    except Exception:
        date_clean = pd.NaT

    return date_clean, time_clean

# ============================================================
# Judge
# ============================================================

JUDGE_PATTERNS = {
    "Ankur Srivastava": [
        "ankur srivastava",
    ],
    "Antara Nath Rivera": [
        "antara nath rivera",
    ],
    "Shauna L. Boliker": [
        "shauna l. boliker",
        "shauna boliker",
        "associate judge shauna l. boliker",
    ],
    "Luciano Pacini Jr.": [
        "luciano pacini jr",
        "luciano pacini jr.",
        "luciano pacini, jr.",
    ],
    "D'Anthony (Tony) Thedford": [
        "judge d'anthony (tony) thedford",
        "d'anthony thedford",
        "tony thedford",
        "thedford",
    ],
    "James Costello": [
        "james costello",
        "james a. costello",
        "james. a costello",
    ],
    "Rivanda Doss Beal": [
        "rivanda doss beal",
        "rivanda b. doss",
        "doss, rivanda",
    ],
    "James Murphy III": [
        "james murphy iii",
        "james p. murphy iii",
        "james murphy v. iii",
        "james murphy",
    ],
    "Deirdre M. Dyer": [
        "deidre dyer",
        "deidre m dyer",
        "deirdre m dyer",
        "deidredyer",
    ],
    "John Hock": [
        "john hock",
    ],
}


def clean_judge(val):
    return match_patterns(val, JUDGE_PATTERNS)


# ============================================================
# Ability to hear
# ============================================================

def match_prefix_patterns(val, patterns, default=np.nan):

    text = normalize_text(val)

    if is_null_value(text):
        return default

    for clean_value, prefixes in patterns.items():
        if text.startswith(tuple(prefixes)):
            return clean_value

    return default

ABILITY_TO_HEAR_PATTERNS = {
    "Yes": ["yes", "good"],
    "No": ["no"],
}

def clean_ability_to_hear(val):

    result = match_patterns(val, ABILITY_TO_HEAR_PATTERNS)

    # no match case
    if result is np.nan:
        return "Sometimes"

    return result

# ============================================================
# Courtroom type
# ============================================================

COURTROOM_TYPE_PATTERNS = {
    "Fishbowl court": ["glass"],
    "Open court": ["open"],
}

def clean_courtroom_type(val):

    result = match_patterns(val, COURTROOM_TYPE_PATTERNS)

    # no match case
    if result is np.nan:
        return "Sometimes"

    return result

COURTROOM_WINDOWS_PATTERNS = {
    "Yes": ["yes"],
    "No": ["no"],
}

def clean_courtroom_has_windows(val):

    result = match_prefix_patterns(val, COURTROOM_WINDOWS_PATTERNS)

    # no match case
    if result is np.nan:
        return "Sometimes"

    return result
# ============================================================
# Prosecutor
# ============================================================

PROSECUTOR_PATTERNS = {
    "Alex Lutzow": [
        "charles alexander lutzow",
        'charles alexander "alex" lutzow iii',
        "alex lutzow",
    ],
    "Ebony Shanklin": [
        "ebony shanklin",
        "ebony shenkin",
    ],
    "John Kyle": [
        "john martin kyle",
        "john kyle",
    ],
    "Ryan Anderson": [
        "ryan anderson",
    ],
    "Adam Weiner": [
        "adam weiner",
        "state, adam weiner",
    ],
    "David Weiner": [
        "david weiner",
    ],
    "Deja Fox": [
        "deja fox",
    ],
}

DEFENSE_ATTORNEY_PATTERNS = {
    "Olivia Dolan": [
        "olivia dolan",
    ],
    "Kameron Clay": [
        "kameron clay",
        "kam clay",
    ],
    "Michael Bramer": [
        "michael bramer",
    ],
    "Kayla Fox": [
        "kayla fox",
    ],
    "Carolyn Paul": [
        "carolyn paul",
    ],
    "Jessica Becker": [
        "jessica becker",
        "jessica becker.",
    ],
    "Erica Tietz": [
        "erica tietz",
    ],
    "Wendy Fawcett": [
        "wendy fawcett",
    ],
    "Sean Craig": [
        "sean craig",
    ],
    "Lillian Parker": [
        "lillian parker",
    ],
    "Claire Bullington": [
        "claire bullington",
        "clare bullington",
    ],
    "Manpreet Chauhan": [
        "manpreet chauhan",
    ],
    "Jennifer Wagner": [
        "jennifer wagner",
    ],
    "Isra Rahman": [
        "isra rahman",
    ],
    "Amber Klinge": [
        "amber klinge",
    ],
    "Julian Sanchez Crozier": [
        "julian sanchez crozier",
    ],
    "Christopher Dallas": [
        "christopher dallas",
    ],
    "Michael Thompson": [
        "michael thompson",
    ],
    "Dawn Ewing": [
        "dawn ewing",
    ],
    "Faith Le": [
        "faith le",
    ],
    "Morghan Gleason": [
        "morghan gleason",
    ],
    "Alana De Leon": [
        "alana de leon",
    ],
    "Saundra Gavazzi": [
        "saundra gavazzi",
    ],
    "Andrew Goldberg": [
        "andrew goldberg",
    ],
    "Paul Burnson": [
        "paul burnson",
    ],
}


def clean_prosecutor(val):
    return match_patterns(val, PROSECUTOR_PATTERNS)

def clean_defense_attorney(val):

    matches = match_patterns(val, DEFENSE_ATTORNEY_PATTERNS)

    if not matches:
        return np.nan

    return " | ".join(matches)

DEFENSE_ATTORNEY_TYPE_PATTERNS = {
    "Public Defender": [
        "public defender",
        "public defenders office",
        "asst public defender",
        "asst. public defender",
        "assistant public defender",
        "law student working with the public defender",
        "711 through public defender",
        "711 law student working with the public defender",
        "senior law student through public defender",
    ],
    "Private Attorney": [
        "private attorney",
        "private for",
        "private counsel",
        "private attorney representing",
    ],
}

def clean_defense_attorney_type(val):
    return match_patterns(val, DEFENSE_ATTORNEY_TYPE_PATTERNS)

def clean_accused_gender_presentation(val):

    text = normalize_text(val)

    if is_null_value(text):
        return "Unknown"

    if "masculine" in text:
        return "Masculine"

    if "feminine" in text:
        return "Feminine"

    return "Unknown"


def clean_accused_perceived_race(val):

    text = normalize_text(val)

    if is_null_value(text):
        return "Unknown"

    # standardize variant spelling first
    text = text.replace("latine", "latinx")

    # match in priority order (safe + explicit)
    if "black" in text:
        return "Black/AA"

    if "latinx" in text:
        return "Latinx/Hispanic"

    if "white" in text:
        return "White"

    if "asian" in text:
        return "Asian"

    return "Unknown"


def extract_pretrial_value(val):
    """
    Returns:
    - number (int)
    - "monitoring"
    - np.nan
    - "X with monitoring"
    """

    text = normalize_text(val)

    if is_null_value(text):
        return np.nan

    # detect monitoring
    is_monitoring = "monitor" in text

    # extract first number found
    nums = re.findall(r"\d+", text)

    if nums:
        num = int(nums[0])

        if is_monitoring:
            return f"{num} with monitoring"
        return num

    if is_monitoring:
        return "monitoring"

    return np.nan


def extract_pretrial_components(val):

    text = normalize_text(val)

    if is_null_value(text):
        return np.nan, False

    # detect monitoring
    is_monitoring = "monitor" in text

    # extract number
    nums = re.findall(r"\d+", text)

    if nums:
        return int(nums[0]), is_monitoring

    if is_monitoring:
        return np.nan, True

    return np.nan, False

def split_pretrial_column(df, col):

    values = df[col].apply(extract_pretrial_components)

    df[f"{col}_clean__value"] = values.apply(lambda x: x[0])
    df[f"{col}_clean__is_monitoring"] = values.apply(lambda x: x[1])

    return df

def clean_pretrial_failure_to_appear(val):
    return extract_pretrial_value(val)

def clean_pretrial_new_criminal_activity(val):
    return extract_pretrial_value(val)

def clean_pretrial_supervision_score(val):
    return extract_pretrial_value(val)


PRIMARY_CHARGE_PATTERNS = {
    "UPWF": [
        "upwf",
        "unlawful possession of weapon (firearm)",
        "unlawful possession of weapon",
    ],
    "UPW": [
        "upw",
        "unlawful possession of weapon",
        "upw by a felon",
    ],
    "AUPW": [
        "aup w",
        "aupw",
        "aggravated unlawful possession of weapon",
        "aggravated upw",
    ],
    "Armed robbery": [
        "armed robbery",
    ],
    "Aggravated battery": [
        "aggravated battery",
    ],
    "Aggravated assault": [
        "aggravated assault",
    ],
    "Residential burglary": [
        "residential burglary",
        "residential burgulary",  # typo in your data
    ],
    "Aggravated DUI": [
        "aggravated dui",
    ],
}

def clean_primary_charge(val):
    return match_patterns(val, PRIMARY_CHARGE_PATTERNS)

GUN_FOUND_PATTERNS = {
    "Yes": ["yes"],
    "No": ["no"],
}
def clean_gun_found(val):

    result = match_patterns(val, GUN_FOUND_PATTERNS)

    if pd.isna(result):
        return "Unknown"

    return result

FAMILY_PRESENT_PATTERNS = {
    "Yes": ["yes"],
    "No": ["no"],
}

def clean_family_present(val):

    result = match_patterns(val, FAMILY_PRESENT_PATTERNS)

    if pd.isna(result):
        return "Unknown"

    return result

def extract_dependants_flags(val):

    text = normalize_text(val)

    if is_null_value(text):
        return False, False, False

    # normalize
    text = text.lower()

    # CHILDREN detection
    has_children = (
        "child" in text
        or "children" in text
        or "stepchild" in text
        or "stepchildren" in text
    )

    # JOB detection
    has_job = "job" in text

    # SPOUSE detection
    has_spouse = "spouse" in text or "fiancé" in text or "fiance" in text

    return has_children, has_job, has_spouse

EM_PATTERNS = [
    "electronic monitoring",
    "em ",
    "em,",        # safety for punctuation variants
    "em with",
    "ankle monitor",
    "gps monitor",
    "gps monitoring",
]

def clean_outcome(val):

    text = normalize_text(val)

    if is_null_value(text):
        return np.nan

    if "detained" in text:
        return "Detained"

    if "released" in text or "release" in text:
        return "Released"

    return np.nan

def clean_electronic_monitoring(val):

    text = normalize_text(val)

    if is_null_value(text):
        return np.nan

    for pattern in EM_PATTERNS:
        if pattern in text:
            return "Yes"

    return "No"
# ============================================================
# Apply cleaning
# ============================================================

df["source_file_clean"] = df["source_file"].apply(clean_source_file)

df["hearing_date"] = pd.to_datetime(
    df["hearing_date"],
    errors="coerce",
).dt.date

df["judge_header_clean"] = df["judge_header"].apply(clean_judge)
df["judge_case_clean"] = df["judge_case"].apply(clean_judge)

df["arrived_at_clean"] = df["arrived_at"].apply(extract_datetime_components)
df["left_at_clean"] = df["left_at"].apply(extract_datetime_components)

df["time_hearing_began_clean"] = df["time_hearing_began"].apply(extract_datetime_components)
df["time_hearing_ended_clean"] = df["time_hearing_ended"].apply(extract_datetime_components)

df["ability_to_hear_clean"] = df["ability_to_hear"].apply(clean_ability_to_hear)

df["courtroom_has_windows_clean"] = df["courtroom_has_windows"].apply(clean_courtroom_has_windows)

df["courtroom_type_clean"] = df["courtroom_type"].apply(clean_courtroom_type)

df["prosecutor_clean"] = df["prosecutor"].apply(clean_prosecutor)

df["defense_attorney_clean"] = df["defense_attorney"].apply(clean_defense_attorney)

df["defense_attorney_type_clean"] = df["defense_attorney_type"].apply(clean_defense_attorney_type)

df["accused_gender_presentation_clean"] = df["accused_gender_presentation"].apply(clean_accused_gender_presentation)

df["accused_perceived_race_clean"] = df["accused_perceived_race"].apply(clean_accused_perceived_race)

df = split_pretrial_column(df, "pretrial_supervision_score")

df['pretrial_failure_to_appear_clean'] = df['pretrial_failure_to_appear'].apply(clean_pretrial_failure_to_appear)
df['pretrial_new_criminal_activity_clean'] = df['pretrial_new_criminal_activity'].apply(clean_pretrial_new_criminal_activity)
df['pretrial_supervision_score_clean'] = df['pretrial_supervision_score'].apply(clean_pretrial_supervision_score)

df['primary_charge_clean'] = df['primary_charge'].apply(clean_primary_charge)

df["arrest_datetime_clean__date"] = df["arrest_datetime"].apply(
    lambda x: extract_datetime_components(x)[0]
)

df["arrest_datetime_clean__time"] = df["arrest_datetime"].apply(
    lambda x: extract_datetime_components(x)[1]
)

df[
    [
        "dependants_spouse_job__has_children_clean",
        "dependants_spouse_job__has_job_clean",
        "dependants_spouse_job__has_spouse_clean",
    ]
] = df["dependants_spouse_job"].apply(
    lambda x: pd.Series(extract_dependants_flags(x))
)

df["outcome_clean"] = df["outcome"].apply(clean_outcome)

df["electronic_monitoring_clean"] = df["outcome"].apply(clean_electronic_monitoring)
# ============================================================
# Save
# ============================================================

df.to_csv(
    "output_clean.csv",
    index=False,
)

print("Saved output_clean.csv")