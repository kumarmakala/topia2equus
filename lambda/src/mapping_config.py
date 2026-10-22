"""Mapping rules for the Topia (source) -> Equus EAFH (target) migration.

Everything a functional reviewer may need to change lives here, so the
transformation code in transform.py stays generic.

Source workbook ("FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS"):
  row 2 = Equus technical field (TABLE!FIELD), row 3 = Equus UI name,
  row 4 = Topia field name, data from row 5.
Target workbook ("DATA MIGRATION SE EAFH TEST FILE"):
  row 1 = format instructions, row 2 = field names, row 3 = import headers,
  data from row 4.
"""

SOURCE_SHEET = "Full Historical Past Assignment"
SOURCE_KEY_ROW = 2
SOURCE_UI_ROW = 3
SOURCE_TOPIA_ROW = 4
SOURCE_DATA_START_ROW = 5
SOURCE_FIRST_COL = 2  # column A holds row labels / employee full name

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
    "host job level": "ASSIGNMENT!TO_SALARY_GRADE",  # row 2 wrongly says TO_TITLE
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

# Source column driving the "Cancelled" phase rule (see note on source col DH).
CANCELLATION_REASON_LABEL = "cancellation reason"
CANCELLED_PHASE = "Cancelled"

# Default effective/begin dates to the assignment start date when the source
# has a value but no date of its own.
DEFAULT_DATES_FROM_START = True

# Source placeholders treated as empty.
BLANK_TOKENS = {"not provided", "n/a", "na", "null", "-"}

# Dropdown value translations, keyed by Equus header; lookups are
# case-insensitive. Unlisted values pass through and are listed in the report.
VALUE_MAPS = {
    "EMPLOYEE!TYPE": {"employee": "Employee", "user": "User"},
    "EMPLOYEE!EMPLOYMENT_STATUS": {"employed": "Active"},
}

# Resource dropdowns (coordinator, HRBP, managers...) need Equus resource
# usernames. Fill as {"Topia display name or email": "equus.username"}.
RESOURCE_MAP = {}

# Country name -> ISO 3166-1 alpha-3, for Equus "Match Country Code" columns.
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
    import re
    for name, iso in re.findall(r"\s*(.+?)\|([A-Z]{3})", text):
        table[name.strip().lower()] = iso
    return table


COUNTRY_ISO3 = _parse_countries(_COUNTRIES)
ISO3_CODES = set(COUNTRY_ISO3.values())
