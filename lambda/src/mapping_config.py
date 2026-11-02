"""Mapping rules for the Topia (source) -> Equus EAFH (target) migration.

Everything a functional reviewer may need to change lives here, so the
transformation code in transform.py stays generic.

Source workbook ("FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS"):
  row 2 = Equus technical field (TABLE!FIELD), row 3 = Equus UI name,
  row 4 = Topia field name, data from row 5.
Target workbook ("DATA MIGRATION SE EAFH TEST FILE"):
  row 1 = format instructions, row 2 = field names, row 3 = import headers,
  data from row 4.

Business rules that change values (derivations, value maps, defaulted dates,
column re-pointing, country conversion) are declared with their approval
status in the versioned business_rules.json next to this file and loaded by
load_business_rules(); reference lists live in reference_data/*.json and are
loaded by load_reference_data(). Both loaders reject invalid configuration
with ConfigurationError instead of ignoring it.
"""

import json
import re
from dataclasses import dataclass
from datetime import datetime
from fnmatch import fnmatchcase
from pathlib import Path

SOURCE_SHEET = "Full Historical Past Assignment"
SOURCE_KEY_ROW = 2
SOURCE_UI_ROW = 3
SOURCE_TOPIA_ROW = 4
SOURCE_DATA_START_ROW = 5
SOURCE_FIRST_COL = 2  # column A holds row labels / employee full name

# Business colour coding of source header rows 1-4 (exact ARGB fills).
# red = not required, yellow = pending clarification with the customer,
# cyan = Equus custom field, peach = standard field. Precedence follows the
# order of COLOUR_PRECEDENCE; any other fill is reported as unknown.
SOURCE_COLOUR_ROWS = (1, 2, 3, 4)
# Q8 (Jake, jake-1.xlsx 'datamapping clarrification shef' column J): the
# brown fill has no business meaning, so it is a known colour that decides
# nothing (it is no longer reported as UNKNOWN_COLOUR).
NO_MEANING_COLOUR = "none (Q8: no meaning)"
FILL_COLOURS = {
    "FFFF0000": "red",
    "FFFFFF00": "yellow",
    "FF00FFFF": "cyan",
    "FFFCE5CD": "peach",
    "FFB45F06": NO_MEANING_COLOUR,
}
COLOUR_PRECEDENCE = ("red", "yellow", "cyan", NO_MEANING_COLOUR, "peach")

TARGET_SHEET = "Sheet1"
TARGET_RULE_ROW = 1
TARGET_HEADER_ROW = 3
TARGET_DATA_START_ROW = 4

# Fixed Equus control columns (target row 3 headers).
COL_EMPLOYEE_ACTION = "A"
COL_ASSIGNMENT_COMMAND = "Assignment Command"
COL_EMPLOYEE_NUMBER = "Employee Number"
COL_ASSIGNMENT_NUMBER = "Assignment Number"
COL_ASSIGNMENT_IDENTIFIER = "Assignment Identifier"

# Employee action for the first assignment of an employee in a file, and for
# any further assignments of the same employee.
EMPLOYEE_ACTION_FIRST = "INSERT"
EMPLOYEE_ACTION_REPEAT = "IGNORE"
NEW_ASSIGNMENT_NUMBER = "NEW"

# Topia column (row 4 label) used as the Equus Assignment Identifier, which
# ties the UPDATE child rows to their INSERT row. The Topia Assignment ID is
# unique per assignment; an employee can have several historical assignments.
ASSIGNMENT_IDENTIFIER_LABEL = "assignment id"

# Source row-2 technical fields whose spelling differs from the Equus headers.
# (regex, replacement) applied in order to the upper-cased key.
KEY_ALIASES = [
    (r"^(ASSIGNMENT!(?:FROM|TO)_COMPANY)_LEVEL(\d)$", r"\1_SEGMENT_LEVEL_\2"),
    (r"^ASSIGNMENT!GROUP$", "ASSIGNMENT!GROUP_NAME"),
    (r"^EMPLOYEE!SYSTEM_NUMBER$", COL_EMPLOYEE_NUMBER),
]

# Source columns whose row 2 is a note, a wrong field or a missing field.
# Keyed by the normalised (lower-case, single-spaced) Topia row-4 label, or the
# row-3 UI label when row 4 is empty. Value None = deliberately not migrated.
FIELD_OVERRIDES = {
    "assignment id": "ASSIGNMENT_CF!CF_TEXT029",  # Topia File Number (legacy id)
    "nationality 3": "EMPLOYEE_CF!CF_TEXT013",
    "status category": "ASSIGNMENT!PHASE",
    "home cost center name": "ASSIGNMENT_CF!CF_TEXT023",
    "host cost center name": "ASSIGNMENT_CF!CF_TEXT024",
    "mobility cost center code": "ASSIGNMENT_CF!CF_TEXT027",
    "mobility cost center name": "ASSIGNMENT_CF!CF_TEXT028",
    "expense cost center code": "ASSIGNMENT_CF!CF_TEXT032",
    "expense cost center name": "ASSIGNMENT_CF!CF_TEXT033",
    "host position": "ASSIGNMENT_CF!CF_TEXT025",
    "host employment type": "ASSIGNMENT_CF!CF_TEXT026",
    # "host job level" -> TO_SALARY_GRADE is business rule
    # host_job_level_to_salary_grade (Q32) in business_rules.json.
    "home scheduled weekly hours": "ASSIGNMENT_CF!CF_NUMBER004",
    "end of assignment reason": "ASSIGNMENT_CF!CF_TEXT030",
    "global id": "EMPLOYEE_CF!CF_TEXT014",
    "pre-hire id": "EMPLOYEE_CF!CF_TEXT015",
    "primary reasons for selecting this employee for the role (business justification)":
        "UT_GOVERNANCE_QUESTIONNAIRE2!BUSINESS_JUSTIFICATION",
    "home annualized base pay currency": "ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_CURRENCY",
    "home annualized bse pay amount": "ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_AMOUNT",
    "home profit sharing plan frequency": "UT_CALCULATION_INPUT!HOME_PROFIT_SHARING_PLAN_FREQUENCY",
    "host profit sharing plan frequency": "UT_CALCULATION_INPUT!HOST_PROFIT_SHARING_PLAN_FREQUENCY",
    "host other benefits deduction currency": "UT_CALCULATION_INPUT!EMPLOYEE_ELECTED_BENEFITS_CURRENCY",
    "host other benefits deduction amount": "UT_CALCULATION_INPUT!EMPLOYEE_ELECTED_BENEFITS",
    "home family allowance currency": "UT_CALCULATION_INPUT!FAMILY_ALLOWANCE_CURRENCY",
    "host family size - do we need to calculate based on spouse/dependent at host?":
        "ASSIGNMENT_FAMILY_SIZE!HOST_FAMILY_SIZE",
    "home family size": "ASSIGNMENT_FAMILY_SIZE!HOME_FAMILY_SIZE",
    # Open question in the source sheet ("need to confirm with Anand"), and
    # the candidate target CF expects a country code, not a nationality.
    "pii nationality 1": None,
}

# Address columns -> Equus ASSIGNMENT_MAILING_ADDRESS!LOCATION_TYPE.
# Each non-empty address becomes one child record.
ADDRESS_COLUMNS = {
    "personal home address with effective date": "Departure Residence",
    "personal host address with effective date": "Destination Residence",
    "work host address with effective date": "Destination Office",
    "work home address with effective date": "Departure Office",
}
# The "Location Type - ..." source columns only describe the constants above.
IGNORED_LABEL_PREFIXES = ("location type - ",)

# Contact columns -> (ASSIGNMENT_EMPLOYEE_CONTACT!LOCATION_TYPE, CONTACT_TYPE).
CONTACT_COLUMNS = {
    "phone number": ("Employee Preferred", "Telephone"),
    "mobile phone number": ("Personal Mobile", "Telephone"),
}

# Equus headers that take ISO 3166-1 alpha-3 codes. Listed explicitly because
# the template's row-1 rules are shifted one column against the row-3 keys in
# columns Q-T (EMPLOYEE!BIRTH_COUNTRY has no rule; CF_TEXT016 has one).
COUNTRY_HEADERS = {
    "ASSIGNMENT!FROM_COUNTRY",
    "ASSIGNMENT!TO_COUNTRY",
    "EMPLOYEE!PRIMARY_CITIZENSHIP",
    "EMPLOYEE!SECONDARY_CITIZENSHIP",
    "EMPLOYEE!BIRTH_COUNTRY",
    "ASSIGNMENT_MAILING_ADDRESS!COUNTRY_CODE",
}

# Equus headers a business rule may fill when the mapped source cell is blank
# are derived from the approved rules (BusinessRules.applied_targets()), so
# reconciliation only treats them as explained when the rule was applied.

# Source placeholders treated as empty.
BLANK_TOKENS = {"not provided", "n/a", "na", "null", "-"}
# Placeholders that also appear glued to other text (source N6 holds
# "Not ProvidedHome"). Multi-word tokens only, so real names such as "Nancy"
# are never blanked by a short token like "na".
BLANK_PREFIX_TOKENS = ("not provided",)

# Dropdown value translations are business rule
# value_maps_employee_type_and_status (Q35) in business_rules.json.

# Plausibility ranges (inclusive) for numeric fields. Out-of-range values are
# kept but warned. The 80-hour ceiling is a review threshold, not a business
# rule (source CU6 holds 158 host weekly hours); adjust once confirmed.
NUMERIC_RANGES = {
    "ASSIGNMENT_CF!CF_NUMBER003": (0, 80),  # host scheduled weekly hours
    "ASSIGNMENT_CF!CF_NUMBER004": (0, 80),  # home scheduled weekly hours
}

# (code header, name header) pairs. A name equal to its code, or shaped like a
# code, is warned (source BL5 holds the code "1000-US018907_NVC" as the name).
CODE_NAME_PAIRS = [
    ("ASSIGNMENT_CF!CF_TEXT005", "ASSIGNMENT_CF!CF_TEXT023"),  # home cost center
    ("ASSIGNMENT_CF!CF_TEXT012", "ASSIGNMENT_CF!CF_TEXT024"),  # host cost center
    ("ASSIGNMENT_CF!CF_TEXT027", "ASSIGNMENT_CF!CF_TEXT028"),  # mobility cost center
    ("ASSIGNMENT_CF!CF_TEXT032", "ASSIGNMENT_CF!CF_TEXT033"),  # expense cost center
]
COST_CENTER_CODE_RE = r"^[\w-]+_[A-Z]VC$"  # every code in the sample ends _NVC / _DVC

# Family record tables: spouse/partner (one per assignment, on the INSERT
# row) and dependents (one per row, any number; Q19 ' Mny dependents Ex').
FAMILY_PREFIXES = ("SPOUSE_DOMESTIC_PARTNER", "ASSIGNMENT_DEPENDENT")
# Q10 / Q20 (Jake, jake-1.xlsx 'datamapping clarrification shef' J11, J21):
# a spouse/partner or dependent without names is importable and one populated
# field creates a record. Missing names are a DATA_QUALITY note for business
# review, never a hold, a drop or an invented name.
FAMILY_NAME_FIELDS = ("FIRST_NAME", "LAST_NAME")
FAMILY_NAME_NOTE = "record has no first/last name; importable per Q10/Q20, business review recommended"
# Family dropdowns whose values are checked against the Equus lists
# (ALLOWED_VALUES_FILE) and never converted. Q14 / Q20: lists still to be
# provided by Jake, so without a list every value is reported "not in list".
FAMILY_LIST_HEADERS = (
    "SPOUSE_DOMESTIC_PARTNER!RELATIONSHIP",
    "ASSIGNMENT_DEPENDENT!RELATIONSHIP",
)
FAMILY_LIST_OWNER = ("Q14", "Jake (Equus lists)")

# Dependent exceptions: code -> (Q, owner, meaning). Held dependents are
# never attached to an assignment.
DEPENDENT_HOLDS = {
    "HELD_DEPENDENTS_UNKEYED": ("Q11", "Alka / Topia (keyed dependent extract, D3)",
                                "dependent sheet has no Employee ID + Assignment ID key columns; nothing attached"),
    "HELD_DEPENDENTS_NOT_CONFIGURED": ("Q11", "Alka / Topia (keyed dependent extract, D3)",
                                       "keyed dependent sheet is not the configured dependent_source.sheet; "
                                       "nothing attached"),
    "HELD_DEPENDENTS_GATED": ("Q11", "Alka / Topia (keyed dependent extract, D3)",
                              "dependent_source is not approved_for_import; nothing attached"),
    "HELD_DEPENDENTS_SHEET_MISSING": ("Q11", "Alka / Topia (keyed dependent extract, D3)",
                                      "configured dependent_source.sheet not found in the source workbook"),
    "HELD_DEPENDENTS_NO_ASSIGNMENT_ID": ("Q18", "Alka / Topia (keyed dependent extract)",
                                         "dependent row has no Assignment ID; never linked by Employee Number"),
    "HELD_DEPENDENTS_UNMATCHED": ("Q18", "Alka / Topia (keyed dependent extract)",
                                  "Assignment ID not found among the assignments of this file"),
    "HELD_DEPENDENTS_EMPLOYEE_MISMATCH": ("Q18", "Alka / Topia (keyed dependent extract)",
                                          "Employee ID differs from the Employee Number of the matched assignment"),
    "HELD_DEPENDENTS_AMBIGUOUS": ("Q1/Q18", "Jake (Equus) / Alka, John (Topia column)",
                                  "Assignment ID occurs on more than one assignment row (or the row has no own "
                                  "Assignment ID); dependents not attached"),
}

# Optional Equus dropdown lists, {"EQUUS!HEADER": ["Value", ...]}, bundled
# next to this file when Equus supplies them (Q14). A missing file or a
# header without a list means every value sent to that dropdown is reported
# as unvalidated and counts against import readiness.
ALLOWED_VALUES_FILE = "equus_lists.json"

# Resource dropdowns (coordinator, HRBP, managers...) need Equus resource
# usernames: reference_data/resource_usernames.json (Q2). RESOURCE_MAP below
# is derived from it and stays empty until that list is approved for import.

# Country name -> ISO 3166-1 alpha-3, for Equus "Match Country Code" columns.
# PROVISIONAL (Q12/Q3): hand-written, not the Equus country list. Used only
# through business rule country_name_to_iso3; while that rule is not
# approved_for_import every converted value is reported as a provisional
# conversion and counts against import readiness.
_COUNTRIES = """
Afghanistan|AFG Aland Islands|ALA Albania|ALB Algeria|DZA American Samoa|ASM Andorra|AND Angola|AGO
Anguilla|AIA Antarctica|ATA Antigua and Barbuda|ATG Argentina|ARG Armenia|ARM Aruba|ABW Australia|AUS
Austria|AUT Azerbaijan|AZE Bahamas|BHS Bahrain|BHR Bangladesh|BGD Barbados|BRB Belarus|BLR Belgium|BEL
Belize|BLZ Benin|BEN Bermuda|BMU Bhutan|BTN Bolivia|BOL Bosnia and Herzegovina|BIH Botswana|BWA
Brazil|BRA British Virgin Islands|VGB Brunei|BRN Brunei Darussalam|BRN Bulgaria|BGR Burkina Faso|BFA
Burundi|BDI Cabo Verde|CPV Cape Verde|CPV Cambodia|KHM Cameroon|CMR Canada|CAN Cayman Islands|CYM
Central African Republic|CAF Chad|TCD Chile|CHL China|CHN Colombia|COL Comoros|COM Congo|COG
Democratic Republic of the Congo|COD Congo, Democratic Republic of the|COD Costa Rica|CRI
Cote d'Ivoire|CIV Côte d'Ivoire|CIV Ivory Coast|CIV Croatia|HRV Cuba|CUB Curacao|CUW Curaçao|CUW
Cyprus|CYP Czech Republic|CZE Czechia|CZE Denmark|DNK Djibouti|DJI Dominica|DMA Dominican Republic|DOM
Ecuador|ECU Egypt|EGY El Salvador|SLV Equatorial Guinea|GNQ Eritrea|ERI Estonia|EST Eswatini|SWZ
Swaziland|SWZ Ethiopia|ETH Fiji|FJI Finland|FIN France|FRA French Guiana|GUF French Polynesia|PYF
Gabon|GAB Gambia|GMB Georgia|GEO Germany|DEU Ghana|GHA Gibraltar|GIB Greece|GRC Greenland|GRL
Grenada|GRD Guadeloupe|GLP Guam|GUM Guatemala|GTM Guernsey|GGY Guinea|GIN Guinea-Bissau|GNB Guyana|GUY
Haiti|HTI Honduras|HND Hong Kong|HKG Hungary|HUN Iceland|ISL India|IND Indonesia|IDN Iran|IRN Iraq|IRQ
Ireland|IRL Isle of Man|IMN Israel|ISR Italy|ITA Jamaica|JAM Japan|JPN Jersey|JEY Jordan|JOR
Kazakhstan|KAZ Kenya|KEN Kiribati|KIR Kosovo|XKX Kuwait|KWT Kyrgyzstan|KGZ Laos|LAO
Lao People's Democratic Republic|LAO Latvia|LVA Lebanon|LBN Lesotho|LSO Liberia|LBR Libya|LBY
Liechtenstein|LIE Lithuania|LTU Luxembourg|LUX Macao|MAC Macau|MAC Madagascar|MDG Malawi|MWI
Malaysia|MYS Maldives|MDV Mali|MLI Malta|MLT Marshall Islands|MHL Martinique|MTQ Mauritania|MRT
Mauritius|MUS Mayotte|MYT Mexico|MEX Micronesia|FSM Moldova|MDA Monaco|MCO Mongolia|MNG Montenegro|MNE
Montserrat|MSR Morocco|MAR Mozambique|MOZ Myanmar|MMR Namibia|NAM Nauru|NRU Nepal|NPL Netherlands|NLD
New Caledonia|NCL New Zealand|NZL Nicaragua|NIC Niger|NER Nigeria|NGA North Korea|PRK
North Macedonia|MKD Macedonia|MKD Northern Mariana Islands|MNP Norway|NOR Oman|OMN Pakistan|PAK
Palau|PLW Palestine|PSE Panama|PAN Papua New Guinea|PNG Paraguay|PRY Peru|PER Philippines|PHL
Poland|POL Portugal|PRT Puerto Rico|PRI Qatar|QAT Reunion|REU Réunion|REU Romania|ROU Russia|RUS
Russian Federation|RUS Rwanda|RWA Saint Kitts and Nevis|KNA Saint Lucia|LCA
Saint Vincent and the Grenadines|VCT Samoa|WSM San Marino|SMR Sao Tome and Principe|STP
Saudi Arabia|SAU Senegal|SEN Serbia|SRB Seychelles|SYC Sierra Leone|SLE Singapore|SGP
Sint Maarten|SXM Slovakia|SVK Slovenia|SVN Solomon Islands|SLB Somalia|SOM South Africa|ZAF
South Korea|KOR Korea, Republic of|KOR Republic of Korea|KOR South Sudan|SSD Spain|ESP Sri Lanka|LKA
Sudan|SDN Suriname|SUR Sweden|SWE Switzerland|CHE Syria|SYR Taiwan|TWN Tajikistan|TJK Tanzania|TZA
Thailand|THA Timor-Leste|TLS Togo|TGO Tonga|TON Trinidad and Tobago|TTO Tunisia|TUN Turkey|TUR
Turkiye|TUR Türkiye|TUR Turkmenistan|TKM Turks and Caicos Islands|TCA Tuvalu|TUV Uganda|UGA
Ukraine|UKR United Arab Emirates|ARE UAE|ARE United Kingdom|GBR UK|GBR Great Britain|GBR
United States|USA United States of America|USA US|USA U.S.A.|USA Uruguay|URY Uzbekistan|UZB
Vanuatu|VUT Vatican City|VAT Holy See|VAT Venezuela|VEN Vietnam|VNM Viet Nam|VNM
US Virgin Islands|VIR Yemen|YEM Zambia|ZMB Zimbabwe|ZWE
"""


def _parse_countries(text):
    table = {}
    # Entries are "Name|ISO"; names contain spaces, so split on the ISO code.
    for name, iso in re.findall(r"\s*(.+?)\|([A-Z]{3})", text):
        table[name.strip().lower()] = iso
    return table


COUNTRY_ISO3 = _parse_countries(_COUNTRIES)
ISO3_CODES = set(COUNTRY_ISO3.values())


# =========================================================================== versioned business configuration
#
# business_rules.json declares every rule below with its Q number, approval
# status and owner. A rule changes import output only while its entry says
# approved_for_import: true; otherwise transform.py leaves the value alone
# (or, for DIAGNOSTIC_RULES, performs the conversion but marks it provisional)
# and reports the rule in the Rules and Readiness sheets.

CONFIG_DIR = Path(__file__).resolve().parent
BUSINESS_RULES_FILE = CONFIG_DIR / "business_rules.json"
REFERENCE_DATA_DIR = CONFIG_DIR / "reference_data"

# Approval vocabulary of clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md.
# A Jake answer is evidence, not approval: approved_for_import stays false
# until the owner explicitly sets it.
RULE_STATUSES = (
    "confirmed_by_jake",
    "partially_clarified",
    "pending_business_approval",
    "pending_equus_configuration",
    "open",
)

RULE_INTERNATIONAL_RELO = "international_relo_derivation"
RULE_CANCELLATION_PHASE = "cancellation_reason_sets_phase"
RULE_DEFAULT_DATES = "default_effective_dates_from_start"
RULE_VALUE_MAPS = "value_maps_employee_type_and_status"
RULE_HOST_JOB_LEVEL = "host_job_level_to_salary_grade"
RULE_COUNTRY_ISO3 = "country_name_to_iso3"

# Rules that re-point one source column (by normalised Topia label) to an
# Equus header. Not approved = the column is HELD, never mapped by its row 2.
COLUMN_OVERRIDE_RULES = (RULE_HOST_JOB_LEVEL,)
# Rules whose conversion is still performed while not approved, so the
# report can show what it would produce; every result is "provisional".
DIAGNOSTIC_RULES = (RULE_COUNTRY_ISO3,)

# Owner / question of readiness items that are not a single rule. Taken from
# the evidence assessment and OPEN_QUESTIONS.md; not approvals.
READINESS_ITEMS = {
    "pending_column": ("yellow (see OPEN_QUESTIONS.md section 6)", "Alka / John (mapping review)"),
    "missing_target": ("Q6/Q7/Q29", "Jake (Equus template) / Schneider"),
    "unknown_colour": ("Q8", "Alka (source colour legend)"),
    "dropdown": ("Q14", "Jake (Equus lists)"),
    "identifier": ("Q1", "Jake (Equus) / Alka, John (Topia column)"),
    "reconciliation": ("Q27", "Schneider + Jake"),
    "template_configuration": ("Q5/Q29", "Jake (Equus template)"),
}


class ConfigurationError(ValueError):
    """Invalid or unsupported configuration. `source` is the file and `path`
    the JSON path of the offending element (e.g. $.rules[2].status)."""

    def __init__(self, source, path: str, message: str):
        self.source = str(source)
        self.path = path
        self.message = message
        super().__init__(f"{source}: {path}: {message}")


class _Invalid(Exception):
    def __init__(self, sub: str, msg: str):
        super().__init__(msg)
        self.sub, self.msg = sub, msg


class _DuplicateKey(ValueError):
    pass


_Q_RE = re.compile(r"^Q\d{1,2}(/Q\d{1,2})*$")
_HEADER_RE = re.compile(r"^[A-Z0-9_]+![A-Z0-9_]+$")


def _text(v, sub=""):
    if not (isinstance(v, str) and v.strip()):
        raise _Invalid(sub, "must be a non-empty string")


def _label(v, sub=""):
    _text(v, sub)
    if v != " ".join(v.split()).lower():
        raise _Invalid(sub, "must be a normalised (lower-case, single-spaced) Topia label")


def _header(v, sub=""):
    if not (isinstance(v, str) and _HEADER_RE.match(v)):
        raise _Invalid(sub, f"must be an Equus header like TABLE!FIELD, got {v!r}")


def _header_list(v, sub=""):
    if not isinstance(v, list) or not v:
        raise _Invalid(sub, "must be a non-empty list of Equus headers")
    for i, h in enumerate(v):
        _header(h, f"{sub}[{i}]")
    if len(set(v)) != len(v):
        raise _Invalid(sub, "contains duplicate headers")


def _dates_map(v, sub=""):
    if not isinstance(v, dict) or not v:
        raise _Invalid(sub, "must be a non-empty object {date header: [trigger headers]}")
    for date_header, triggers in v.items():
        _header(date_header, f"{sub}.{date_header}")
        _header_list(triggers, f"{sub}.{date_header}")


def _value_maps(v, sub=""):
    if not isinstance(v, dict) or not v:
        raise _Invalid(sub, "must be a non-empty object {Equus header: {source value: Equus value}}")
    for header, mapping in v.items():
        _header(header, f"{sub}.{header}")
        if not isinstance(mapping, dict) or not mapping:
            raise _Invalid(f"{sub}.{header}", "must be a non-empty object")
        for src, dst in mapping.items():
            if not src or src != src.strip().lower():
                raise _Invalid(f"{sub}.{header}.{src}",
                               "source values are matched case-insensitively and must be lower-case")
            _text(dst, f"{sub}.{header}.{src}")


def _one_of(*allowed):
    def check(v, sub=""):
        if v not in allowed:
            raise _Invalid(sub, f"must be one of {list(allowed)}")
    return check


def _iso_timestamp(v, sub=""):
    _text(v, sub)
    try:
        datetime.fromisoformat(v)
    except ValueError:
        raise _Invalid(sub, "must be an ISO 8601 date or timestamp") from None


def _q_number(v, sub=""):
    if not (isinstance(v, str) and _Q_RE.match(v)):
        raise _Invalid(sub, "must be a question number like 'Q12' or 'Q22/Q23'")


def _status(v, sub=""):
    if v not in RULE_STATUSES:
        raise _Invalid(sub, f"unknown status {v!r}; allowed: {', '.join(RULE_STATUSES)}")


def _boolean(v, sub=""):
    if not isinstance(v, bool):
        raise _Invalid(sub, "must be true or false")


# Parameter schema of every rule the engine implements: exactly these keys.
RULE_PARAMETERS = {
    RULE_INTERNATIONAL_RELO: {
        "target": _header, "home_country": _header, "host_country": _header,
        "value_if_different": _text, "value_if_same": _text,
    },
    RULE_CANCELLATION_PHASE: {"source_label": _label, "target": _header, "value": _text},
    RULE_DEFAULT_DATES: {"start_header": _header, "effective_dates": _dates_map, "child_dates": _header_list},
    RULE_VALUE_MAPS: {"maps": _value_maps},
    RULE_HOST_JOB_LEVEL: {"source_label": _label, "target": _header},
    RULE_COUNTRY_ISO3: {"table": _one_of("mapping_config.COUNTRY_ISO3")},
}

_META_CHECKS = (("version", _text), ("effective_version", _text), ("generated_at", _iso_timestamp))
_TOP_REQUIRED = ("version", "effective_version", "generated_at", "rules")
_TOP_OPTIONAL = ("description", "dependent_source")
_RULE_REQUIRED = ("id", "Q", "title", "status", "approved_for_import", "owner", "parameters")
_RULE_OPTIONAL = ("notes", "approval_reference")


def _run(check, value, source, path: str) -> None:
    try:
        check(value, "")
    except _Invalid as exc:
        raise ConfigurationError(source, path + exc.sub, exc.msg) from None


def _check_object(obj, path: str, source, required, optional) -> None:
    if not isinstance(obj, dict):
        raise ConfigurationError(source, path, "must be a JSON object")
    for key in obj:
        if key not in required and key not in optional:
            raise ConfigurationError(source, f"{path}.{key}", "unknown key")
    for key in required:
        if key not in obj:
            raise ConfigurationError(source, f"{path}.{key}", "missing required key")


def _strict_pairs(pairs):
    obj = {}
    for key, value in pairs:
        if key in obj:
            raise _DuplicateKey(key)
        obj[key] = value
    return obj


def _reject_constant(name):
    raise ValueError(f"non-standard JSON constant {name}")


def _read_json(path: Path):
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ConfigurationError(path, "$", f"cannot read file ({exc.__class__.__name__})") from None
    try:
        return json.loads(text, object_pairs_hook=_strict_pairs, parse_constant=_reject_constant)
    except _DuplicateKey as exc:
        raise ConfigurationError(path, "$", f"duplicate key {exc.args[0]!r}") from None
    except ValueError as exc:
        raise ConfigurationError(path, "$", f"invalid JSON: {exc}") from None


def _walk_headers(value, path: str):
    if isinstance(value, str):
        if _HEADER_RE.match(value):
            yield path, value
    elif isinstance(value, list):
        for i, item in enumerate(value):
            yield from _walk_headers(item, f"{path}[{i}]")
    elif isinstance(value, dict):
        for key, item in value.items():
            if _HEADER_RE.match(key):
                yield f"{path}.{key}", key
            yield from _walk_headers(item, f"{path}.{key}")


@dataclass(frozen=True)
class BusinessRule:
    id: str
    q: str
    title: str
    status: str
    approved_for_import: bool
    owner: str
    parameters: dict
    path: str
    notes: str = ""
    approval_reference: str = ""


@dataclass(frozen=True)
class DependentSource:
    """Keyed Topia dependent sheet (Q11/Q18/Q19). Dependents are attached
    only while `sheet` is set and approved_for_import is true, and only by
    Assignment ID (never by Employee Number alone). `version` is the version
    of the business_rules.json that declares the section."""
    version: str
    q: str
    status: str
    approved_for_import: bool
    owner: str
    sheet: str | None
    header_row: int
    data_start_row: int
    employee_label: str
    assignment_label: str
    child_groups: str
    notes: str = ""

    @property
    def configured(self) -> bool:
        return self.sheet is not None

    @property
    def enabled(self) -> bool:
        return self.configured and self.approved_for_import


DEPENDENT_CHILD_GROUPS = ("repeated_column_groups",)
_DEP_REQUIRED = ("Q", "status", "approved_for_import", "owner", "parameters")
_DEP_OPTIONAL = ("notes", "approval_reference")
_DEP_PARAMS = ("sheet", "header_row", "data_start_row", "key_columns", "child_groups")


def _positive_int(v, sub=""):
    if not isinstance(v, int) or isinstance(v, bool) or v < 1:
        raise _Invalid(sub, "must be a positive integer (worksheet row number)")


def _load_dependent_source(raw, source, version: str) -> DependentSource:
    p = "$.dependent_source"
    _check_object(raw, p, source, _DEP_REQUIRED, _DEP_OPTIONAL)
    _run(_q_number, raw["Q"], source, f"{p}.Q")
    _run(_status, raw["status"], source, f"{p}.status")
    _run(_boolean, raw["approved_for_import"], source, f"{p}.approved_for_import")
    _run(_text, raw["owner"], source, f"{p}.owner")
    for key in _DEP_OPTIONAL:
        if key in raw:
            _run(_text, raw[key], source, f"{p}.{key}")
    params = raw["parameters"]
    pp = f"{p}.parameters"
    _check_object(params, pp, source, _DEP_PARAMS, ())
    if params["sheet"] is not None:
        _run(_text, params["sheet"], source, f"{pp}.sheet")
        if params["sheet"] == SOURCE_SHEET:
            raise ConfigurationError(source, f"{pp}.sheet", "must not be the assignment sheet SOURCE_SHEET")
    _run(_positive_int, params["header_row"], source, f"{pp}.header_row")
    _run(_positive_int, params["data_start_row"], source, f"{pp}.data_start_row")
    if params["data_start_row"] <= params["header_row"]:
        raise ConfigurationError(source, f"{pp}.data_start_row", "must be greater than header_row")
    keys = params["key_columns"]
    _check_object(keys, f"{pp}.key_columns", source, ("employee", "assignment"), ())
    _run(_label, keys["employee"], source, f"{pp}.key_columns.employee")
    _run(_label, keys["assignment"], source, f"{pp}.key_columns.assignment")
    if keys["employee"] == keys["assignment"]:
        raise ConfigurationError(source, f"{pp}.key_columns",
                                 "employee and assignment key columns must differ (Assignment ID is the join key)")
    _run(_one_of(*DEPENDENT_CHILD_GROUPS), params["child_groups"], source, f"{pp}.child_groups")
    if raw["approved_for_import"] and params["sheet"] is None:
        raise ConfigurationError(source, f"{p}.approved_for_import",
                                 "cannot be true while parameters.sheet is null")
    return DependentSource(
        version=version, q=raw["Q"], status=raw["status"], approved_for_import=raw["approved_for_import"],
        owner=raw["owner"], sheet=params["sheet"], header_row=params["header_row"],
        data_start_row=params["data_start_row"], employee_label=keys["employee"],
        assignment_label=keys["assignment"], child_groups=params["child_groups"], notes=raw.get("notes", ""))


@dataclass(frozen=True)
class BusinessRules:
    source: str
    version: str
    effective_version: str
    generated_at: str
    rules: dict  # rule id -> BusinessRule, in file order
    dependent_source: DependentSource | None = None  # None = section absent (no keyed source)

    @property
    def file_name(self) -> str:
        return Path(self.source).name

    def rule(self, rule_id: str) -> BusinessRule:
        try:
            return self.rules[rule_id]
        except KeyError:
            raise ConfigurationError(self.source, "$.rules", f"rule {rule_id!r} is not declared") from None

    def rule_enabled(self, rule_id: str) -> bool:
        return self.rule(rule_id).approved_for_import

    def params(self, rule_id: str) -> dict:
        return self.rule(rule_id).parameters

    def header_refs(self):
        """(JSON path, Equus header) for every header a rule refers to."""
        for rule in self.rules.values():
            yield from _walk_headers(rule.parameters, f"{rule.path}.parameters")

    def applied_targets(self) -> set:
        """Headers an *approved* rule fills when the mapped source cell is blank."""
        targets = set()
        for rule_id in (RULE_INTERNATIONAL_RELO, RULE_CANCELLATION_PHASE):
            if self.rule_enabled(rule_id):
                targets.add(self.params(rule_id)["target"])
        if self.rule_enabled(RULE_DEFAULT_DATES):
            targets.update(self.params(RULE_DEFAULT_DATES)["effective_dates"])
        return targets


def load_business_rules(path=None) -> BusinessRules:
    """Load and strictly validate business_rules.json (or `path`)."""
    source = Path(path) if path is not None else BUSINESS_RULES_FILE
    data = _read_json(source)
    _check_object(data, "$", source, _TOP_REQUIRED, _TOP_OPTIONAL)
    for key, check in _META_CHECKS:
        _run(check, data[key], source, f"$.{key}")
    if "description" in data:
        _run(_text, data["description"], source, "$.description")
    raw_rules = data["rules"]
    if not isinstance(raw_rules, list) or not raw_rules:
        raise ConfigurationError(source, "$.rules", "must be a non-empty list")

    rules: dict = {}
    for i, raw in enumerate(raw_rules):
        p = f"$.rules[{i}]"
        _check_object(raw, p, source, _RULE_REQUIRED, _RULE_OPTIONAL)
        _run(_text, raw["id"], source, f"{p}.id")
        rule_id = raw["id"]
        if rule_id in rules:
            raise ConfigurationError(source, f"{p}.id",
                                     f"duplicate rule id {rule_id!r} (first declared at {rules[rule_id].path})")
        if rule_id not in RULE_PARAMETERS:
            raise ConfigurationError(source, f"{p}.id",
                                     f"unsupported rule id {rule_id!r}; supported: {', '.join(RULE_PARAMETERS)}")
        _run(_q_number, raw["Q"], source, f"{p}.Q")
        _run(_text, raw["title"], source, f"{p}.title")
        _run(_status, raw["status"], source, f"{p}.status")
        _run(_boolean, raw["approved_for_import"], source, f"{p}.approved_for_import")
        _run(_text, raw["owner"], source, f"{p}.owner")
        for key in _RULE_OPTIONAL:
            if key in raw:
                _run(_text, raw[key], source, f"{p}.{key}")
        schema = RULE_PARAMETERS[rule_id]
        _check_object(raw["parameters"], f"{p}.parameters", source, tuple(schema), ())
        for name, check in schema.items():
            _run(check, raw["parameters"][name], source, f"{p}.parameters.{name}")
        rules[rule_id] = BusinessRule(
            id=rule_id, q=raw["Q"], title=raw["title"], status=raw["status"],
            approved_for_import=raw["approved_for_import"], owner=raw["owner"],
            parameters=raw["parameters"], path=p, notes=raw.get("notes", ""),
            approval_reference=raw.get("approval_reference", ""),
        )
    missing = [rule_id for rule_id in RULE_PARAMETERS if rule_id not in rules]
    if missing:
        raise ConfigurationError(source, "$.rules", f"missing rule(s) {', '.join(missing)}: every rule the "
                                                    "engine implements must be declared with its approval status")
    dep = _load_dependent_source(data["dependent_source"], source, data["version"]) if "dependent_source" in data else None
    return BusinessRules(source=str(source), version=data["version"], effective_version=data["effective_version"],
                         generated_at=data["generated_at"], rules=rules, dependent_source=dep)


# --------------------------------------------------------------------------- reference data
#
# One JSON per list in reference_data/. Absent file, empty values or
# approved_for_import false = every value of the covered headers is an
# exception ("unknown value"), never a guess. applies_to uses fnmatch
# patterns over Equus headers.
#   kind "lookup":          source value -> Equus value (values: {str: str})
#   kind "allowed_values":  converted value must be a key of `values`

COST_CENTRE_CODE_HEADERS = tuple(code for code, _ in CODE_NAME_PAIRS)

REFERENCE_LISTS = {
    "resource_usernames": {"kind": "lookup", "value_type": str, "Q": "Q2", "owner": "Jake (Equus)",
                           "applies_to": ("*RESOURCE*",)},
    "cost_centres": {"kind": "allowed_values", "value_type": str, "Q": "Q15", "owner": "Jake (Equus) / Alka (Topia)",
                     "applies_to": COST_CENTRE_CODE_HEADERS},
    "country_state": {"kind": "allowed_values", "value_type": list, "Q": "Q12/Q3", "owner": "Jake (Equus)",
                      "applies_to": tuple(sorted(COUNTRY_HEADERS))},
}
RESOURCE_LIST_ID = "resource_usernames"

_REF_REQUIRED = ("list_id", "version", "effective_version", "generated_at", "Q", "title", "status",
                 "approved_for_import", "owner", "values")
_REF_OPTIONAL = ("description", "notes", "approval_reference")


@dataclass(frozen=True)
class ReferenceList:
    list_id: str
    kind: str
    q: str
    title: str
    status: str
    approved_for_import: bool
    owner: str
    version: str
    effective_version: str
    applies_to: tuple
    values: dict
    source: str
    present: bool = True

    @property
    def usable(self) -> bool:
        return self.present and self.approved_for_import and bool(self.values)

    @property
    def state(self) -> str:
        if not self.present:
            return "file absent"
        if not self.values:
            return "empty"
        if not self.approved_for_import:
            return "not approved_for_import"
        return "approved"

    def applies(self, header: str) -> bool:
        return any(fnmatchcase(header, pattern) for pattern in self.applies_to)


def _check_values(values, value_type, source, path: str) -> None:
    if not isinstance(values, dict):
        raise ConfigurationError(source, path, "must be a JSON object")
    for key, value in values.items():
        if not key.strip():
            raise ConfigurationError(source, f"{path}.{key}", "keys must be non-empty")
        if value_type is str:
            _run(_text, value, source, f"{path}.{key}")
        elif not isinstance(value, list) or not all(isinstance(v, str) and v.strip() for v in value):
            raise ConfigurationError(source, f"{path}.{key}", "must be a list of non-empty strings")


def load_reference_data(directory=None) -> dict:
    """Load and strictly validate reference_data/*.json (or `directory`)."""
    directory = Path(directory) if directory is not None else REFERENCE_DATA_DIR
    if directory.exists() and not directory.is_dir():
        raise ConfigurationError(directory, "$", "reference data location is not a directory")
    files = {p.stem: p for p in sorted(directory.glob("*.json"))} if directory.is_dir() else {}
    for stem, path in files.items():
        if stem not in REFERENCE_LISTS:
            raise ConfigurationError(path, "$", f"unsupported reference list {stem!r}; "
                                                f"supported: {', '.join(REFERENCE_LISTS)}")
    out = {}
    for list_id, decl in REFERENCE_LISTS.items():
        path = files.get(list_id)
        if path is None:
            out[list_id] = ReferenceList(
                list_id=list_id, kind=decl["kind"], q=decl["Q"], title=f"{list_id} (file absent)", status="open",
                approved_for_import=False, owner=decl["owner"], version="", effective_version="",
                applies_to=decl["applies_to"], values={}, source=str(directory / f"{list_id}.json"), present=False)
            continue
        data = _read_json(path)
        _check_object(data, "$", path, _REF_REQUIRED, _REF_OPTIONAL)
        if data["list_id"] != list_id:
            raise ConfigurationError(path, "$.list_id", f"must equal the file name {list_id!r}")
        for key, check in _META_CHECKS:
            _run(check, data[key], path, f"$.{key}")
        _run(_q_number, data["Q"], path, "$.Q")
        _run(_text, data["title"], path, "$.title")
        _run(_status, data["status"], path, "$.status")
        _run(_boolean, data["approved_for_import"], path, "$.approved_for_import")
        _run(_text, data["owner"], path, "$.owner")
        for key in _REF_OPTIONAL:
            if key in data:
                _run(_text, data[key], path, f"$.{key}")
        _check_values(data["values"], decl["value_type"], path, "$.values")
        out[list_id] = ReferenceList(
            list_id=list_id, kind=decl["kind"], q=data["Q"], title=data["title"], status=data["status"],
            approved_for_import=data["approved_for_import"], owner=data["owner"], version=data["version"],
            effective_version=data["effective_version"], applies_to=decl["applies_to"], values=data["values"],
            source=str(path))
    return out


# --------------------------------------------------------------------------- shipped configuration
# Loaded at import so a broken shipped file fails the Lambda cold start
# instead of producing output. transform() reloads (or takes a test copy).

BUSINESS_RULES = load_business_rules()
REFERENCE_DATA = load_reference_data()


def rule_enabled(rule_id: str, rules: BusinessRules | None = None) -> bool:
    """True only when the rule's configuration entry is approved_for_import."""
    return (rules if rules is not None else BUSINESS_RULES).rule_enabled(rule_id)


# Read-only views for local/build_field_mapping.py (documentation output).
VALUE_MAPS = BUSINESS_RULES.params(RULE_VALUE_MAPS)["maps"]
CANCELLATION_REASON_LABEL = BUSINESS_RULES.params(RULE_CANCELLATION_PHASE)["source_label"]
CANCELLED_PHASE = BUSINESS_RULES.params(RULE_CANCELLATION_PHASE)["value"]
# Keyed dependent source in force (business_rules.json $.dependent_source).
DEPENDENT_SOURCE = BUSINESS_RULES.dependent_source
RESOURCE_MAP = dict(REFERENCE_DATA[RESOURCE_LIST_ID].values) if REFERENCE_DATA[RESOURCE_LIST_ID].usable else {}
