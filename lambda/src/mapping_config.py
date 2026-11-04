"""Central mapping rules for the Topia (source) -> Equus EAFH (target) migration.

Everything a functional reviewer may need to change lives here, so the
transformation code in transform.py stays generic.

To stop migrating a Topia column later, add its normalised name to
EXCLUDE_FIELDS (one line) - no code change required.
"""


SOURCE_SHEET = "Topia Source"
SOURCE_HEADER_ROW = 1
SOURCE_DATA_START_ROW = 2

TARGET_SHEET = " Mny dependents Ex"  # NB: leading space, as in the Equus file
TARGET_RULE_ROW = 1
TARGET_UI_ROW = 2
TARGET_HEADER_ROW = 3           # Equus DB fields live here
TARGET_DATA_START_ROW = 4


# Fixed Equus control columns (exact row-3 header text in the template).
COL_EMPLOYEE_COMMAND = "Employee Command"
COL_ASSIGNMENT_COMMAND = "Assignment Command"
COL_EMPLOYEE_NUMBER = "Employee Number"
COL_ASSIGNMENT_NUMBER = "Assignment Number"
COL_ASSIGNMENT_IDENTIFIER = "Assignment Identifier"

EMPLOYEE_COMMAND_FIRST = "INSERT"   
EMPLOYEE_COMMAND_REPEAT = "IGNORE"  
ASSIGNMENT_COMMAND_INSERT = "INSERT"
ASSIGNMENT_COMMAND_UPDATE = "UPDATE"
NEW_ASSIGNMENT_NUMBER = "NEW"


# Normalised Topia header supplying each identity key. Change in one place.
EMPLOYEE_NUMBER_SOURCE = "employee id"          # Talent Link ID -> Employee Number
ASSIGNMENT_IDENTIFIER_SOURCE = "topia id"       # legacy Topia ID -> Assignment Identifier
# (also written to ASSIGNMENT_CF!CF_TEXT029 via FIELD_MAP below.)


# Key aliases (COMPANY_LEVELn -> COMPANY_SEGMENT_LEVEL_n, GROUP -> GROUP_NAME)
# and all approved overrides are already baked in.
FIELD_MAP = {
    'alternate email': 'EMPLOYEE_CF!CF_TEXT008',
    'arrival date': 'ASSIGNMENT!HOST_COUNTRY_ARRIVAL_DATE',
    'assignment end date': 'ASSIGNMENT!ACTUAL_END_DATE',
    'assignment start date': 'ASSIGNMENT!START_DATE',
    'assignment type': 'ASSIGNMENT!TYPE',
    'assignment type category': 'ASSIGNMENT!RELOCATION_POLICY_ID',
    'business title': 'ASSIGNMENT_CF!CF_TEXT014',
    'citizenship 1': 'EMPLOYEE!PRIMARY_CITIZENSHIP',
    'citizenship 2': 'EMPLOYEE!SECONDARY_CITIZENSHIP',
    'country of birth': 'EMPLOYEE!BIRTH_COUNTRY',
    'date of birth': 'EMPLOYEE!BIRTH_DATE',
    'departure date': 'ASSIGNMENT!HOME_COUNTRY_DEPARTURE_DATE',
    'email': 'EMPLOYEE!USER_ID_EMAIL',
    'employee benefits manager - host': 'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE13_ID',
    'employment status': 'EMPLOYEE!EMPLOYMENT_STATUS',
    'end of assignment reason': 'ASSIGNMENT_CF!CF_TEXT030',
    'expense cost center code': 'ASSIGNMENT_CF!CF_TEXT032',
    'expense cost center name': 'ASSIGNMENT_CF!CF_TEXT033',
    'first name': 'EMPLOYEE!FIRST_NAME',
    'gender': 'EMPLOYEE!GENDER',
    'global id': 'EMPLOYEE_CF!CF_TEXT014',
    'group name': 'ASSIGNMENT!GROUP_NAME',
    'home annual vacation days': 'ASSIGNMENT_COMPENSATION_POLICY!VACATION_DAYS_PER_YEAR',
    'home annualized base pay currency': 'ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_CURRENCY',
    'home annualized bse pay amount': 'ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_AMOUNT',
    'home business line': 'ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_3',
    'home city': 'ASSIGNMENT!FROM_CITY',
    'home cost center code': 'ASSIGNMENT_CF!CF_TEXT005',
    'home cost center name': 'ASSIGNMENT_CF!CF_TEXT023',
    'home country': 'ASSIGNMENT!FROM_COUNTRY',
    'home division': 'ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_4',
    'home employment type': 'ASSIGNMENT_CF!CF_TEXT007',
    'home entity': 'ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_1',
    'home family allowance currency': 'UT_CALCULATION_INPUT!FAMILY_ALLOWANCE_CURRENCY',
    'home hfm code': 'ASSIGNMENT_CF!CF_TEXT004',
    'home job level': 'ASSIGNMENT!FROM_SALARY_GRADE',
    'home job profile': 'ASSIGNMENT_CF!CF_TEXT002',
    'home job title': 'ASSIGNMENT!FROM_TITLE',
    'home office location': 'ASSIGNMENT_CF!CF_TEXT006',
    'home owner status': 'ASSIGNMENT!HOMEOWNER',
    'home pay installments': 'ASSIGNMENT_CF!CF_NUMBER001',
    'home position': 'ASSIGNMENT_CF!CF_TEXT008',
    'home profit sharing plan frequency': 'UT_CALCULATION_INPUT!HOME_PROFIT_SHARING_PLAN_FREQUENCY',
    'home role region': 'ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_2',
    'home scheduled weekly hours': 'ASSIGNMENT_CF!CF_NUMBER004',
    'home state': 'ASSIGNMENT!FROM_STATE_PROVINCE',
    'home target bonus %': 'ASSIGNMENT_COMPENSATION_POLICY_CF!CF_NUMBER001',
    'host business line': 'ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_3',
    'host city': 'ASSIGNMENT!TO_CITY',
    'host cost center code': 'ASSIGNMENT_CF!CF_TEXT012',
    'host cost center name': 'ASSIGNMENT_CF!CF_TEXT024',
    'host country': 'ASSIGNMENT!TO_COUNTRY',
    'host division': 'ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_4',
    'host employee id': 'EMPLOYEE_CF!CF_TEXT001',
    'host employment type': 'ASSIGNMENT_CF!CF_TEXT026',
    'host entity': 'ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_1',
    'host job level': 'ASSIGNMENT!TO_SALARY_GRADE',
    'host job profile': 'ASSIGNMENT_CF!CF_TEXT009',
    'host job title': 'ASSIGNMENT!TO_TITLE',
    'host office location': 'ASSIGNMENT_CF!CF_TEXT013',
    'host other benefits deduction amount': 'UT_CALCULATION_INPUT!EMPLOYEE_ELECTED_BENEFITS',
    'host other benefits deduction currency': 'UT_CALCULATION_INPUT!EMPLOYEE_ELECTED_BENEFITS_CURRENCY',
    'host pay installments': 'ASSIGNMENT_CF!CF_NUMBER002',
    'host position': 'ASSIGNMENT_CF!CF_TEXT025',
    'host profit sharing plan frequency': 'UT_CALCULATION_INPUT!HOST_PROFIT_SHARING_PLAN_FREQUENCY',
    'host role region': 'ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_2',
    'host scheduled weekly hours': 'ASSIGNMENT_CF!CF_NUMBER003',
    'host state': 'ASSIGNMENT!TO_STATE_PROVINCE',
    'host supervisory organization': 'ASSIGNMENT_CF!CF_TEXT011',
    'host target bonus %': 'ASSIGNMENT_COMPENSATION_POLICY_CF!CF_NUMBER002',
    'hr advisor home name': 'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE1_ID',
    'hr advisor host name': 'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE2_ID',
    'hr leader host (host hr svp)': 'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE6_ID',
    'last name': 'EMPLOYEE!LAST_NAME',
    'line manager home name': 'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE3_ID',
    'line manager host name': 'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE4_ID',
    'marital status': 'EMPLOYEE!MARITAL_STATUS',
    'middle name': 'EMPLOYEE!MIDDLE_NAME',
    'mobility cost center code': 'ASSIGNMENT_CF!CF_TEXT027',
    'mobility cost center name': 'ASSIGNMENT_CF!CF_TEXT028',
    'mobility driver': 'ASSIGNMENT!RELOCATION_REASON',
    'mobility host contact (imp)': 'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE10_ID',
    'move initiator name': 'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE11_ID',
    'nationality 1': 'EMPLOYEE_CF!CF_TEXT002',
    'nationality 2': 'EMPLOYEE_CF!CF_TEXT005',
    'nationality 3': 'EMPLOYEE_CF!CF_TEXT013',
    'original hire date': 'EMPLOYEE!HIRE_DATE',
    'person status': 'EMPLOYEE!TYPE',
    'pre-hire id': 'EMPLOYEE_CF!CF_TEXT015',
    'primary mobility contact': 'ASSIGNMENT!COORDINATOR_RESOURCE_ASSIGNMENT_ID',
    'primary reasons for selecting this employee for the role (business justification)': 'UT_GOVERNANCE_QUESTIONNAIRE2!BUSINESS_JUSTIFICATION',
    'relocation contact': 'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE12_ID',
    'spouse at host': 'SPOUSE_DOMESTIC_PARTNER!STATUS',
    'spouse date of birth': 'SPOUSE_DOMESTIC_PARTNER!BIRTH_DATE',
    'spouse first name': 'SPOUSE_DOMESTIC_PARTNER!FIRST_NAME',
    'spouse gender': 'SPOUSE_DOMESTIC_PARTNER!GENDER',
    'spouse last name': 'SPOUSE_DOMESTIC_PARTNER!LAST_NAME',
    'spouse relationship': 'SPOUSE_DOMESTIC_PARTNER!RELATIONSHIP',
    'status category': 'ASSIGNMENT!PHASE',
    'talent profile': 'UT_GOVERNANCE_QUESTIONNAIRE2!TALENT_RATING',
    'topia id': 'ASSIGNMENT_CF!CF_TEXT029',
    'vip flag': 'ASSIGNMENT!VIP',
}


RELABEL_HEADERS_FROM_REFERENCE = True   
RELABEL_DB_FIELDS = True                


HEADER_MAP = {
    'EMPLOYEE!FIRST_NAME': ('EMPLOYEE!FIRST_NAME', 'First Name', None),
    'EMPLOYEE!LAST_NAME': ('EMPLOYEE!LAST_NAME', 'Last Name', None),
    'EMPLOYEE!MIDDLE_NAME': ('EMPLOYEE!MIDDLE_NAME', 'Middle Name', None),
    'EMPLOYEE!USER_ID_EMAIL': ('EMPLOYEE!USER_ID_EMAIL', 'Email Address', None),
    'EMPLOYEE_CF!CF_TEXT008': ('EMPLOYEE_CF!CF_TEXT008', 'Alternate Email Address', None),
    'EMPLOYEE!GENDER': ('EMPLOYEE!GENDER', 'Gender', 'Match dropdown List'),
    'EMPLOYEE!BIRTH_DATE': ('EMPLOYEE!BIRTH_DATE', 'Birth Date', 'US Date Format'),
    'EMPLOYEE!MARITAL_STATUS': ('EMPLOYEE!MARITAL_STATUS', 'Marital Status', None),
    'EMPLOYEE_CF!CF_TEXT002': ('EMPLOYEE_CF!CF_TEXT002', 'Primary Nationality', None),
    'EMPLOYEE_CF!CF_TEXT005': ('EMPLOYEE_CF!CF_TEXT005', 'Secondary Nationality', None),
    'EMPLOYEE!PRIMARY_CITIZENSHIP': ('EMPLOYEE!PRIMARY_CITIZENSHIP', 'Primary Citizenship', 'Match dropdown List'),
    'EMPLOYEE!SECONDARY_CITIZENSHIP': ('EMPLOYEE!SECONDARY_CITIZENSHIP', 'Secondary Citizenship', 'Match dropdown List'),
    'EMPLOYEE!BIRTH_COUNTRY': ('EMPLOYEE!BIRTH_COUNTRY', 'Country of Birth', 'Match dropdown List'),
    'EMPLOYEE!TYPE': ('EMPLOYEE!TYPE', None, 'Match dropdown list'),
    'ASSIGNMENT_CF!CF_TEXT014': ('ASSIGNMENT_CF!CF_TEXT014', 'Career Level', None),
    'EMPLOYEE!EMPLOYMENT_STATUS': ('EMPLOYEE!EMPLOYMENT_STATUS', 'Employment Status', 'Match dropdown List'),
    'ASSIGNMENT!TYPE': ('ASSIGNMENT!TYPE', 'Mobility Type', 'Match dropdown List'),
    'ASSIGNMENT!RELOCATION_POLICY_ID': ('ASSIGNMENT!RELOCATION_POLICY_ID', 'Mobility Policy', 'Match dropdown List'),
    'ASSIGNMENT_CF!CF_NUMBER001': ('ASSIGNMENT_CF!CF_NUMBER001', 'Number of payroll installment - Home', None),
    'ASSIGNMENT_CF!CF_NUMBER002': ('ASSIGNMENT_CF!CF_NUMBER002', 'Number of payroll installment - Host', None),
    'ASSIGNMENT!START_DATE': ('ASSIGNMENT!START_DATE', 'Start Date', 'US Date Format'),
    'ASSIGNMENT!ACTUAL_END_DATE': ('ASSIGNMENT!ACTUAL_END_DATE', 'Actual End Date', 'US Date Format'),
    'ASSIGNMENT!HOME_COUNTRY_DEPARTURE_DATE': ('ASSIGNMENT!HOME_COUNTRY_DEPARTURE_DATE', 'Home Country Departure Date', 'US Date Format'),
    'ASSIGNMENT!HOST_COUNTRY_ARRIVAL_DATE': ('ASSIGNMENT!HOST_COUNTRY_ARRIVAL_DATE', 'Host Country Arrival Date', 'US Date Format'),
    'ASSIGNMENT!FROM_COUNTRY': ('ASSIGNMENT!FROM_COUNTRY', 'From Country', 'Match dropdown List - Country Code'),
    'ASSIGNMENT!FROM_CITY': ('ASSIGNMENT!FROM_CITY', 'From City', None),
    'ASSIGNMENT!FROM_STATE_PROVINCE': ('ASSIGNMENT!FROM_STATE_PROVINCE', 'From State/Province', 'Match dropdown List - State Name'),
    'ASSIGNMENT!TO_COUNTRY': ('ASSIGNMENT!TO_COUNTRY', 'To Country', 'Match dropdown List - Country Code'),
    'ASSIGNMENT!TO_CITY': ('ASSIGNMENT!TO_CITY', 'To City', None),
    'ASSIGNMENT!TO_STATE_PROVINCE': ('ASSIGNMENT!TO_STATE_PROVINCE', 'To State/Province', 'Match dropdown List - State Name'),
    'ASSIGNMENT_CF!CF_TEXT005': ('ASSIGNMENT_CF!CF_TEXT005', 'Home Cost Center', 'Match dropdown list'),
    'ASSIGNMENT_CF!CF_TEXT012': ('ASSIGNMENT_CF!CF_TEXT012', 'Host Cost Center', 'Match dropdown list'),
    'ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_1': ('ASSIGNMENT!FROM_COMPANY_LEVEL1', 'From Legal Entity', 'Match dropdown list'),
    'ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_2': ('ASSIGNMENT!FROM_COMPANY_LEVEL2', 'From Region', 'Match dropdown list'),
    'ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_3': ('ASSIGNMENT!FROM_COMPANY_LEVEL3', 'From Business Unit', 'Match dropdown list'),
    'ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_4': ('ASSIGNMENT!FROM_COMPANY_LEVEL4', 'From Department', 'Match dropdown list'),
    'ASSIGNMENT!FROM_TITLE': ('ASSIGNMENT!FROM_TITLE', 'Home Job Title', 'Match dropdown list'),
    'ASSIGNMENT!FROM_SALARY_GRADE': ('ASSIGNMENT!FROM_SALARY_GRADE', 'Home Job Grade', 'Match dropdown list'),
    'ASSIGNMENT_CF!CF_TEXT002': ('ASSIGNMENT_CF!CF_TEXT002', 'Home Job Profile', None),
    'ASSIGNMENT_CF!CF_TEXT008': ('ASSIGNMENT_CF!CF_TEXT008', 'Home Position Time Type', None),
    'ASSIGNMENT_CF!CF_TEXT007': ('ASSIGNMENT_CF!CF_TEXT007', 'Home Job Level', None),
    'ASSIGNMENT_CF!CF_TEXT006': ('ASSIGNMENT_CF!CF_TEXT006', 'Home Office Location', None),
    'ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_1': ('ASSIGNMENT!TO_COMPANY_LEVEL1', 'To Legal Entity', 'Match dropdown list'),
    'ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_2': ('ASSIGNMENT!TO_COMPANY_LEVEL2', 'To Region', 'Match dropdown list'),
    'ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_3': ('ASSIGNMENT!TO_COMPANY_LEVEL3', 'To Business Unit', 'Match dropdown list'),
    'ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_4': ('ASSIGNMENT!TO_COMPANY_LEVEL4', 'To Department', 'Match dropdown list'),
    'ASSIGNMENT!TO_TITLE': ('ASSIGNMENT!TO_TITLE', 'Host Job Title', 'Match dropdown list'),
    'ASSIGNMENT_CF!CF_TEXT009': ('ASSIGNMENT_CF!CF_TEXT009', 'Host Job Profile', None),
    'ASSIGNMENT_CF!CF_NUMBER003': ('ASSIGNMENT_CF!CF_NUMBER003', 'Host Working Hours', None),
    'ASSIGNMENT_CF!CF_TEXT013': ('ASSIGNMENT_CF!CF_TEXT013', 'Host Office Location', None),
    'ASSIGNMENT_CF!CF_TEXT011': ('ASSIGNMENT_CF!CF_TEXT011', 'Host HFM Code', 'Match dropdown list'),
    'ASSIGNMENT!VIP': ('ASSIGNMENT!VIP', 'VIP', 'Checkbox'),
    'EMPLOYEE_CF!CF_TEXT001': ('EMPLOYEE_CF!CF_TEXT001', 'SESA ID', None),
    'ASSIGNMENT!HOMEOWNER': ('ASSIGNMENT!HOMEOWNER', 'Home Owner Field', 'Yes/No'),
    'ASSIGNMENT!COORDINATOR_RESOURCE_ASSIGNMENT_ID': ('ASSIGNMENT!COORDINATOR_RESOURCE_ASSIGNMENT_ID', 'Mobility Coordinator', 'Match Resource dropdown list'),
    'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE1_ID': ('ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE1_ID', 'Home HRBP', 'Match Resource dropdown'),
    'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE2_ID': ('ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE2_ID', 'Host HRBP', 'Match Resource dropdown'),
    'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE3_ID': ('ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE3_ID', 'Home Manager?', None),
    'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE4_ID': ('ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE4_ID', 'Host Manager?', None),
    'ASSIGNMENT!GROUP_NAME': ('ASSIGNMENT!GROUP', 'Assignment: Group', None),
    'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE11_ID': ('ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE11_ID', 'Move Initiator', 'Match Resource dropdown'),
    'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE12_ID': ('ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE12_ID', 'Onsite Consultant', None),
    'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE10_ID': ('ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE10_ID', 'Host IMP Resource', 'Match Resource dropdown'),
    'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE6_ID': ('ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE6_ID', 'Resource', 'Match Resource dropdown'),
    'ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE13_ID': ('ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE13_ID', 'Executive Rewards', None),
    'ASSIGNMENT_MAILING_ADDRESS!LOCATION_TYPE': ('ASSIGNMENT_MAILING_ADDRESS!LOCATION_TYPE', None, None),
    'ASSIGNMENT_EMPLOYEE_CONTACT!NUMBER_ADDRESS': ('ASSIGNMENT_EMPLOYEE_CONTACT!NUMBER_ADDRESS', 'Employee Contact', None),
    'UT_GOVERNANCE_QUESTIONNAIRE2!TALENT_RATING': ('UT_GOVERNANCE_QUESTIONNAIRE2!TALENT_RATING', None, None),
    'UT_CALCULATION_INPUT!EMPLOYEE_ELECTED_BENEFITS': ('UT_CALCULATION_INPUT!EMPLOYEE_ELECTED_BENEFITS', 'Cost Estimate Additional Inputs Screen', None),
    'ASSIGNMENT_COMPENSATION_POLICY_CF!CF_NUMBER001': ('ASSIGNMENT_COMPENSATION_POLICY_CF!CF_NUMBER001', 'Home STIP %', None),
    'ASSIGNMENT_COMPENSATION_POLICY_CF!CF_NUMBER002': ('ASSIGNMENT_COMPENSATION_POLICY_CF!CF_NUMBER002', 'Host STIP %', None),
    'UT_CALCULATION_INPUT!FAMILY_ALLOWANCE': ('UT_CALCULATION_INPUT!FAMILY_ALLOWANCE', 'Cost Estimate Additional Inputs', None),
    'ASSIGNMENT_CF!CF_TEXT004': ('ASSIGNMENT_CF!CF_TEXT004', 'Assignment CF', 'Match dropdown list'),
    'ASSIGNMENT_COMPENSATION_POLICY!VACATION_DAYS_PER_YEAR': ('ASSIGNMENT_COMPENSATION_POLICY!VACATION_DAYS_PER_YEAR', 'Home Annual Vacation Days', None),
    'ASSIGNMENT!RELOCATION_REASON': ('ASSIGNMENT!RELOCATION_REASON', 'Mobility Reason', 'Match dropdown'),
    'EMPLOYEE!HIRE_DATE': ('EMPLOYEE!HIRE_DATE', 'Original Hire Date', None),
    'ASSIGNMENT!DEFAULT_CURRENCY': ('ASSIGNMENT!DEFAULT_CURRENCY', 'Default Currency', None),
    'UT_CALCULATION_INPUT!RMC': ('UT_CALCULATION_INPUT!RMC', 'RMC they use', None),
    'SPOUSE_DOMESTIC_PARTNER!LAST_NAME': ('SPOUSE_DOMESTIC_PARTNER!LAST_NAME', None, None),
    'SPOUSE_DOMESTIC_PARTNER!FIRST_NAME': ('SPOUSE_DOMESTIC_PARTNER!FIRST_NAME', 'Spouse/Domestic Partner', None),
    'SPOUSE_DOMESTIC_PARTNER!RELATIONSHIP': ('SPOUSE_DOMESTIC_PARTNER!RELATIONSHIP', None, None),
    'SPOUSE_DOMESTIC_PARTNER!STATUS': ('SPOUSE_DOMESTIC_PARTNER!STATUS', None, None),
    'SPOUSE_DOMESTIC_PARTNER!GENDER': ('SPOUSE_DOMESTIC_PARTNER!GENDER', None, None),
    'SPOUSE_DOMESTIC_PARTNER!BIRTH_DATE': ('SPOUSE_DOMESTIC_PARTNER!BIRTH_DATE', None, None),
}

# Normalised Topia headers to NEVER migrate, even if present in FIELD_MAP.
# Add a line here to drop a column in future - no other code change needed.
EXCLUDE_FIELDS = set()

# Known source-header aliases for different Topia extracts / schema versions:
# {alternate normalised header -> canonical normalised header in FIELD_MAP}.
# Exact, documented renames only - NO fuzzy matching.
SOURCE_ALIASES = {
    "date of birth (dob)": "date of birth",
    "e-mail": "email",
    "employee number": "employee id",
    "talent link id": "employee id",
}


# Address text columns -> ASSIGNMENT_MAILING_ADDRESS!LOCATION_TYPE constant.
# Each non-empty address becomes one UPDATE child row.
ADDRESS_COLUMNS = {
    "personal home address with effective date": "Departure Residence",
    "personal host address with effective date": "Destination Residence",
    "work host address with effective date": "Destination Office",
    "work home address with effective date": "Departure Office",
}
# The "Location Type - ..." source columns only describe the constants above.
IGNORED_LABEL_PREFIXES = ("location type - ",)

# Contact columns -> (LOCATION_TYPE, CONTACT_TYPE). Each becomes an UPDATE row.
CONTACT_COLUMNS = {
    "phone number": ("Employee Preferred", "Telephone"),
    "mobile phone number": ("Personal Mobile", "Telephone"),
}

# Dependent columns follow the "Dependent Import" layout: repeated blocks named
# "Child N <field>". Mapping keys match the trailing <field> (normalised).
DEPENDENT_FIELD_MAP = {
    "first name": "ASSIGNMENT_DEPENDENT!FIRST_NAME",
    "last name": "ASSIGNMENT_DEPENDENT!LAST_NAME",
    "gender": "ASSIGNMENT_DEPENDENT!GENDER",
    "date of birth": "ASSIGNMENT_DEPENDENT!BIRTH_DATE",
    "relationship": "ASSIGNMENT_DEPENDENT!RELATIONSHIP",
    "at host": "ASSIGNMENT_DEPENDENT!STATUS",
    "supported - cross cultural training": "ASSIGNMENT_DEPENDENT_CF!CF_TEXT001",
    "supported - child in third party (housing)": "ASSIGNMENT_DEPENDENT_CF!CF_TEXT003",
    "supported - relocation allowance (eligible non-accompanying)": "ASSIGNMENT_DEPENDENT_CF!CF_TEXT006",
}
# Matches "Child 1 First Name", "Dependent 2 Gender", etc. Group 1 = child key.
import re as _re
DEPENDENT_HEADER_RE = _re.compile(r"^(?:child|dependent)\s*(\d+)\s+(.*)$")


# Non-empty Cancellation Reason -> ASSIGNMENT!PHASE = "Cancelled" (Clarification /
# reference note: SE do not import a cancellation date, only the phase).
CANCELLATION_REASON_LABEL = "cancellation reason"
CANCELLED_PHASE = "Cancelled"

# Derive ASSIGNMENT!INTERNATIONAL_RELO from home vs host country when absent.
DERIVE_INTERNATIONAL_RELO = True

# Default effective/begin dates to the assignment start date when a value exists
# but has no date of its own (salary effective date, address begin, contact start).
DEFAULT_DATES_FROM_START = True

# Source placeholders treated as empty (case-insensitive).
BLANK_TOKENS = {"not provided", "n/a", "na", "null", "-", ""}

# Dropdown value translations, keyed by Equus header; lookups case-insensitive.
# Unlisted values pass through and are listed in the mapping report.
VALUE_MAPS = {
    "EMPLOYEE!TYPE": {"employee": "Employee", "user": "User"},
    "EMPLOYEE!EMPLOYMENT_STATUS": {"employed": "Active"},
}

# Resource dropdowns (coordinator, HRBP, managers...) need Equus usernames.
# Fill as {"Topia display name or email": "equus.username"} once provided
# (datamapping-clarification Q2). Unmatched values pass through + are reported.
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
    for name, iso in _re.findall(r"\s*(.+?)\|([A-Z]{3})", text):
        table[name.strip().lower()] = iso
    return table


COUNTRY_ISO3 = _parse_countries(_COUNTRIES)
ISO3_CODES = set(COUNTRY_ISO3.values())
