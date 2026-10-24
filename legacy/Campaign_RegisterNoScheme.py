import streamlit as st
import pandas as pd
from io import BytesIO
import re
import zipfile

# -----------------------------------
# Streamlit Configuration
# -----------------------------------

st.set_page_config(
    page_title="Customer Location Finder",
    page_icon="📍",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📍 Customer State & Language Finder")
st.caption("Upload a customer report plus optional Bss Receipt / Sales Purchase files. "
           "Unknown customers are resolved sequentially: Bss Receipt → Sales Purchase. "
           "Export includes 10 sheets with Mobileno, State, District, Language, Template. "
           "**Rows with empty/blank phone numbers are automatically removed.**")


# -----------------------------------
# Sheet names — tab ≤ 31 chars (Excel limit), full title in A1
# -----------------------------------

SHEET_META = [
    ("1_All_App_Cus",                       "1_All_App_Cus"),
    ("2_Zero_Invested_App_Cus",             "2_Zero_Invested_App_Cus"),
    ("3_Purchase_Matched_App_Reg",          "3_Purchase_Cus_Matched_With_App_Register"),
    ("4_Bss_Matched_App_Reg",               "4_Bss_Cus_Matched_With_App_Register"),
    ("5_Purchase_Unmatched_App",            "5_Purchase_Cus_Unmatched_With_App"),
    ("6_Bss_Unmatched_With_App",            "6_Bss_Cus_Unmatched_With_App"),
    ("7_Purch_Matched_Invested",            "7_Purchase_Cus_Matched_With_App_Invested"),
    ("8_Purch_Matched_No_Invested",         "8_Purchase_Cus_Matched_With_App_No_Invested"),
    ("9_Bss_Matched_Invested",              "9_Bss_Cus_Matched_With_App_Invested"),
    ("10_Bss_Matched_No_Invested",          "10_Bss_Cus_Matched_With_App_No_Invested"),
]


# -----------------------------------
# State Language Mapping
# -----------------------------------

state_language = {
    "Tamil Nadu": "Tamil", "Karnataka": "Kannada", "Kerala": "Malayalam",
    "Andhra Pradesh": "Telugu", "Telangana": "Telugu", "Maharashtra": "Marathi",
    "Gujarat": "Gujarati", "West Bengal": "Bengali", "Odisha": "Odia",
    "Punjab": "Punjabi", "Rajasthan": "Hindi", "Uttar Pradesh": "Hindi",
    "Bihar": "Hindi", "Madhya Pradesh": "Hindi", "Delhi": "Hindi",
    "Haryana": "Hindi", "Jharkhand": "Hindi", "Chhattisgarh": "Hindi",
    "Assam": "Assamese", "Goa": "Konkani", "Manipur": "Manipuri",
    "Mizoram": "Mizo", "Sikkim": "Nepali", "Tripura": "Bengali",
    "Puducherry": "Tamil", "Pondicherry": "Tamil",
    "Jammu & Kashmir": "Urdu/Hindi", "Jammu and Kashmir": "Urdu/Hindi",
    "Ladakh": "Ladakhi", "Himachal Pradesh": "Hindi", "Uttarakhand": "Hindi",
}


state_variations = {
    "tamilnadu": "Tamil Nadu", "tamil nadu": "Tamil Nadu", "tamilnad": "Tamil Nadu",
    "tamailnadu": "Tamil Nadu", "tanilnadu": "Tamil Nadu", "thamilnadu": "Tamil Nadu",
    "thamil nadu": "Tamil Nadu", "thamiznadu": "Tamil Nadu", "thamizh nadu": "Tamil Nadu",
    "thamizhnadu": "Tamil Nadu", "tamizhnadu": "Tamil Nadu", "tamizh nadu": "Tamil Nadu",
    "taminadu": "Tamil Nadu", "taminad": "Tamil Nadu", "tamnad": "Tamil Nadu",
    "tamnadu": "Tamil Nadu", "tami nadu": "Tamil Nadu", "thamil natu": "Tamil Nadu",
    "tamil natu": "Tamil Nadu", "tn": "Tamil Nadu", "tamil": "Tamil Nadu",
    "madras": "Tamil Nadu", "chennai": "Tamil Nadu",
    "karnataka": "Karnataka", "karnatka": "Karnataka", "karnatak": "Karnataka",
    "ka": "Karnataka", "bangalore": "Karnataka", "bengaluru": "Karnataka",
    "kerala": "Kerala", "kerla": "Kerala", "kerela": "Kerala", "keral": "Kerala",
    "kerlaa": "Kerala", "keralam": "Kerala", "kl": "Kerala",
    "kochi": "Kerala", "trivandrum": "Kerala",
    "andhra pradesh": "Andhra Pradesh", "andhra": "Andhra Pradesh",
    "andhrapradesh": "Andhra Pradesh", "ap": "Andhra Pradesh",
    "telangana": "Telangana", "telangna": "Telangana", "ts": "Telangana",
    "hyderabad": "Telangana",
    "maharashtra": "Maharashtra", "maharastra": "Maharashtra",
    "maharashstra": "Maharashtra", "mh": "Maharashtra",
    "mumbai": "Maharashtra", "pune": "Maharashtra",
    "gujarat": "Gujarat", "gujrat": "Gujarat", "gj": "Gujarat", "ahmedabad": "Gujarat",
    "west bengal": "West Bengal", "westbengal": "West Bengal",
    "wb": "West Bengal", "kolkata": "West Bengal",
    "odisha": "Odisha", "orissa": "Odisha", "orrisa": "Odisha", "od": "Odisha",
    "uttar pradesh": "Uttar Pradesh", "uttarpradesh": "Uttar Pradesh",
    "up": "Uttar Pradesh", "lucknow": "Uttar Pradesh",
    "bihar": "Bihar", "br": "Bihar", "patna": "Bihar",
    "rajasthan": "Rajasthan", "rj": "Rajasthan", "jaipur": "Rajasthan",
    "madhya pradesh": "Madhya Pradesh", "madhyapradesh": "Madhya Pradesh",
    "mp": "Madhya Pradesh", "bhopal": "Madhya Pradesh",
    "delhi": "Delhi", "dl": "Delhi",
    "punjab": "Punjab", "pb": "Punjab", "chandigarh": "Punjab",
    "haryana": "Haryana", "haryan": "Haryana", "hr": "Haryana",
    "jharkhand": "Jharkhand", "jh": "Jharkhand",
    "chhattisgarh": "Chhattisgarh", "chattisgarh": "Chhattisgarh", "cg": "Chhattisgarh",
    "assam": "Assam", "as": "Assam", "goa": "Goa", "ga": "Goa",
    "puducherry": "Puducherry", "pondicherry": "Puducherry", "py": "Puducherry",
    "jammu & kashmir": "Jammu & Kashmir", "jammu and kashmir": "Jammu & Kashmir",
    "jammu&kashmir": "Jammu & Kashmir", "jammukashmir": "Jammu & Kashmir",
    "j&k": "Jammu & Kashmir", "jk": "Jammu & Kashmir",
    "jammu": "Jammu & Kashmir", "kashmir": "Jammu & Kashmir",
    "himachal pradesh": "Himachal Pradesh", "himachalpradesh": "Himachal Pradesh",
    "hp": "Himachal Pradesh", "himachal": "Himachal Pradesh",
    "uttarakhand": "Uttarakhand", "uttaranchal": "Uttarakhand", "uk": "Uttarakhand",
}


city_state = {
    "ariyalur": "Tamil Nadu", "chengalpattu": "Tamil Nadu", "chennai": "Tamil Nadu",
    "coimbatore": "Tamil Nadu", "coimpatore": "Tamil Nadu", "coimbtore": "Tamil Nadu",
    "coimbatur": "Tamil Nadu", "coimbathore": "Tamil Nadu", "coyambatore": "Tamil Nadu",
    "coyambathore": "Tamil Nadu", "kovai": "Tamil Nadu", "cbe": "Tamil Nadu",
    "cuddalore": "Tamil Nadu", "dharmapuri": "Tamil Nadu", "dindigul": "Tamil Nadu",
    "erode": "Tamil Nadu", "kallakurichi": "Tamil Nadu", "kanchipuram": "Tamil Nadu",
    "kanyakumari": "Tamil Nadu", "kanniyakumari": "Tamil Nadu",
    "karur": "Tamil Nadu", "krishnagiri": "Tamil Nadu",
    "madurai": "Tamil Nadu", "mayiladuthurai": "Tamil Nadu", "nagapattinam": "Tamil Nadu",
    "namakkal": "Tamil Nadu", "nilgiris": "Tamil Nadu", "perambalur": "Tamil Nadu",
    "pudukkottai": "Tamil Nadu", "ramanathapuram": "Tamil Nadu", "ranipet": "Tamil Nadu",
    "salem": "Tamil Nadu", "sivaganga": "Tamil Nadu", "tenkasi": "Tamil Nadu",
    "thanjavur": "Tamil Nadu", "theni": "Tamil Nadu", "thoothukudi": "Tamil Nadu",
    "tuticorin": "Tamil Nadu", "tiruchirappalli": "Tamil Nadu", "trichy": "Tamil Nadu",
    "tirunelveli": "Tamil Nadu", "tirupathur": "Tamil Nadu", "tiruppur": "Tamil Nadu",
    "tirupur": "Tamil Nadu", "tiruvallur": "Tamil Nadu", "tiruvannamalai": "Tamil Nadu",
    "tiruvarur": "Tamil Nadu", "vellore": "Tamil Nadu", "viluppuram": "Tamil Nadu",
    "virudhunagar": "Tamil Nadu", "karaikudi": "Tamil Nadu", "ooty": "Tamil Nadu",
    "hosur": "Tamil Nadu", "sivakasi": "Tamil Nadu", "rajapalayam": "Tamil Nadu",
    "nagercoil": "Tamil Nadu", "marthandam": "Tamil Nadu", "nagerkovil": "Tamil Nadu",
    "tambaram": "Tamil Nadu", "avadi": "Tamil Nadu", "pallavaram": "Tamil Nadu",
    "poonamallee": "Tamil Nadu", "ambattur": "Tamil Nadu", "porur": "Tamil Nadu",
    "mylapore": "Tamil Nadu", "adyar": "Tamil Nadu",

    "kochi": "Kerala", "cochin": "Kerala", "trivandrum": "Kerala",
    "thiruvananthapuram": "Kerala", "kozhikode": "Kerala", "calicut": "Kerala",
    "thrissur": "Kerala", "trichur": "Kerala", "kollam": "Kerala",
    "alappuzha": "Kerala", "alleppey": "Kerala", "palakkad": "Kerala",
    "palghat": "Kerala", "kannur": "Kerala", "kasargod": "Kerala",
    "pathanamthitta": "Kerala", "idukki": "Kerala", "wayanad": "Kerala",
    "kottayam": "Kerala", "ernakulam": "Kerala", "malappuram": "Kerala",

    "puducherry": "Puducherry", "pondicherry": "Puducherry", "karaikal": "Puducherry",

    "bangalore": "Karnataka", "bengaluru": "Karnataka", "mysore": "Karnataka",
    "mysuru": "Karnataka", "mangalore": "Karnataka", "mangaluru": "Karnataka",
    "hubli": "Karnataka", "dharwad": "Karnataka", "belgaum": "Karnataka",
    "belagavi": "Karnataka", "gulbarga": "Karnataka", "kalaburagi": "Karnataka",
    "bellary": "Karnataka", "ballari": "Karnataka", "tumkur": "Karnataka",
    "tumakuru": "Karnataka", "shimoga": "Karnataka", "shivamogga": "Karnataka",
    "davanagere": "Karnataka", "raichur": "Karnataka", "hassan": "Karnataka",
    "kolar": "Karnataka", "mandya": "Karnataka", "udupi": "Karnataka",
    "chitradurga": "Karnataka", "bidar": "Karnataka",

    "vijayawada": "Andhra Pradesh", "visakhapatnam": "Andhra Pradesh",
    "vizag": "Andhra Pradesh", "guntur": "Andhra Pradesh", "nellore": "Andhra Pradesh",
    "kurnool": "Andhra Pradesh", "tirupati": "Andhra Pradesh",
    "rajahmundry": "Andhra Pradesh", "kakinada": "Andhra Pradesh",
    "chittoor": "Andhra Pradesh", "anantapur": "Andhra Pradesh",
    "kadapa": "Andhra Pradesh", "nandyal": "Andhra Pradesh",
    "hyderabad": "Telangana", "secunderabad": "Telangana", "warangal": "Telangana",
    "nizamabad": "Telangana", "karimnagar": "Telangana", "khammam": "Telangana",

    "mumbai": "Maharashtra", "pune": "Maharashtra", "nagpur": "Maharashtra",
    "thane": "Maharashtra", "nashik": "Maharashtra", "aurangabad": "Maharashtra",
    "solapur": "Maharashtra", "kalyan": "Maharashtra", "navi mumbai": "Maharashtra",
    "ahmedabad": "Gujarat", "surat": "Gujarat", "vadodara": "Gujarat",
    "rajkot": "Gujarat", "bhavnagar": "Gujarat", "jamnagar": "Gujarat",
    "junagadh": "Gujarat", "gandhinagar": "Gujarat", "anand": "Gujarat",
    "kolkata": "West Bengal", "howrah": "West Bengal", "durgapur": "West Bengal",
    "siliguri": "West Bengal", "asansol": "West Bengal",

    "delhi": "Delhi", "new delhi": "Delhi",
    "gurugram": "Haryana", "gurgaon": "Haryana", "faridabad": "Haryana",
    "noida": "Uttar Pradesh", "ghaziabad": "Uttar Pradesh",
    "greater noida": "Uttar Pradesh",

    "lucknow": "Uttar Pradesh", "kanpur": "Uttar Pradesh", "varanasi": "Uttar Pradesh",
    "agra": "Uttar Pradesh", "allahabad": "Uttar Pradesh", "prayagraj": "Uttar Pradesh",
    "meerut": "Uttar Pradesh", "aligarh": "Uttar Pradesh", "bareilly": "Uttar Pradesh",
    "patna": "Bihar", "gaya": "Bihar", "bhagalpur": "Bihar", "muzaffarpur": "Bihar",
    "jaipur": "Rajasthan", "jodhpur": "Rajasthan", "udaipur": "Rajasthan",
    "kota": "Rajasthan", "bikaner": "Rajasthan", "ajmer": "Rajasthan",
    "bhopal": "Madhya Pradesh", "indore": "Madhya Pradesh", "gwalior": "Madhya Pradesh",
    "jabalpur": "Madhya Pradesh", "ujjain": "Madhya Pradesh",

    "chandigarh": "Punjab", "amritsar": "Punjab", "ludhiana": "Punjab",
    "jalandhar": "Punjab", "patiala": "Punjab", "mohali": "Punjab",
    "guwahati": "Assam", "dibrugarh": "Assam", "silchar": "Assam",
    "panaji": "Goa", "margao": "Goa", "vasco": "Goa",
    "bhubaneswar": "Odisha", "cuttack": "Odisha", "rourkela": "Odisha",
    "puri": "Odisha", "sambalpur": "Odisha",
    "jammu": "Jammu & Kashmir", "srinagar": "Jammu & Kashmir",
    "shimla": "Himachal Pradesh", "manali": "Himachal Pradesh",
    "dharamshala": "Himachal Pradesh", "solan": "Himachal Pradesh",
    "dehradun": "Uttarakhand", "haridwar": "Uttarakhand", "rishikesh": "Uttarakhand",
    "nainital": "Uttarakhand",
}


district_state = {
    "ariyalur": "Tamil Nadu", "chengalpattu": "Tamil Nadu", "chennai": "Tamil Nadu",
    "coimbatore": "Tamil Nadu", "cuddalore": "Tamil Nadu", "dharmapuri": "Tamil Nadu",
    "dindigul": "Tamil Nadu", "erode": "Tamil Nadu", "kallakurichi": "Tamil Nadu",
    "kanchipuram": "Tamil Nadu", "kanyakumari": "Tamil Nadu", "karur": "Tamil Nadu",
    "krishnagiri": "Tamil Nadu", "madurai": "Tamil Nadu", "mayiladuthurai": "Tamil Nadu",
    "nagapattinam": "Tamil Nadu", "namakkal": "Tamil Nadu", "nilgiris": "Tamil Nadu",
    "perambalur": "Tamil Nadu", "pudukkottai": "Tamil Nadu", "ramanathapuram": "Tamil Nadu",
    "ranipet": "Tamil Nadu", "salem": "Tamil Nadu", "sivaganga": "Tamil Nadu",
    "tenkasi": "Tamil Nadu", "thanjavur": "Tamil Nadu", "theni": "Tamil Nadu",
    "thoothukudi": "Tamil Nadu", "tiruchirappalli": "Tamil Nadu", "tirunelveli": "Tamil Nadu",
    "tirupathur": "Tamil Nadu", "tiruppur": "Tamil Nadu", "tiruvallur": "Tamil Nadu",
    "tiruvannamalai": "Tamil Nadu", "tiruvarur": "Tamil Nadu", "vellore": "Tamil Nadu",
    "viluppuram": "Tamil Nadu", "virudhunagar": "Tamil Nadu",
    "puducherry": "Puducherry", "karaikal": "Puducherry",
    "bengaluru": "Karnataka", "mysuru": "Karnataka", "mangaluru": "Karnataka",
    "thiruvananthapuram": "Kerala", "kochi": "Kerala", "kozhikode": "Kerala",
    "thrissur": "Kerala", "kollam": "Kerala", "kannur": "Kerala",
    "hyderabad": "Telangana", "visakhapatnam": "Andhra Pradesh",
    "vijayawada": "Andhra Pradesh", "guntur": "Andhra Pradesh",
    "mumbai": "Maharashtra", "pune": "Maharashtra", "nagpur": "Maharashtra",
    "ahmedabad": "Gujarat", "surat": "Gujarat",
    "kolkata": "West Bengal", "howrah": "West Bengal",
    "delhi": "Delhi",
    "lucknow": "Uttar Pradesh", "kanpur": "Uttar Pradesh", "varanasi": "Uttar Pradesh",
    "patna": "Bihar", "jaipur": "Rajasthan", "bhopal": "Madhya Pradesh",
    "ludhiana": "Punjab", "amritsar": "Punjab", "guwahati": "Assam",
    "bhubaneswar": "Odisha", "srinagar": "Jammu & Kashmir", "jammu": "Jammu & Kashmir",
    "shimla": "Himachal Pradesh", "dehradun": "Uttarakhand",
}


PINCODE_RANGES = [
    (600000, 643999, "Tamil Nadu"), (605001, 605013, "Puducherry"),
    (607401, 607409, "Puducherry"), (609601, 609609, "Puducherry"),
    (670000, 696999, "Kerala"), (560000, 591999, "Karnataka"),
    (515000, 534999, "Andhra Pradesh"), (500000, 509999, "Telangana"),
    (400000, 449999, "Maharashtra"), (360000, 395999, "Gujarat"),
    (700000, 700999, "West Bengal"), (710000, 744999, "West Bengal"),
    (110000, 110999, "Delhi"), (120000, 139999, "Haryana"),
    (200000, 247999, "Uttar Pradesh"), (264000, 289999, "Uttar Pradesh"),
    (244000, 263999, "Uttarakhand"), (300000, 349999, "Rajasthan"),
    (450000, 499999, "Madhya Pradesh"), (800000, 859999, "Bihar"),
    (751000, 770999, "Odisha"), (180000, 194999, "Jammu & Kashmir"),
    (171000, 177999, "Himachal Pradesh"),
]


# -----------------------------------
# Text helpers
# -----------------------------------

def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = re.sub(r'[^a-z0-9\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def extract_pincode(text):
    if pd.isna(text):
        return None
    match = re.search(r'\b\d{6}\b', str(text))
    return int(match.group()) if match else None


def get_state_from_pincode(pincode):
    if not pincode:
        return None
    best_state, best_width = None, None
    for start, end, state in PINCODE_RANGES:
        if start <= pincode <= end:
            width = end - start
            if best_width is None or width < best_width:
                best_width, best_state = width, state
    return best_state


def get_state_from_variation(text):
    clean = clean_text(text)
    for variation, state in state_variations.items():
        if variation in clean:
            return state
    return None


def fuzzy_state_match(text):
    clean = clean_text(text)
    patterns = [
        (r'tami[nl]?[a]?[d]?[nu]?[au]?[d]?', "Tamil Nadu"),
        (r'tam[il]{1,2}[a]?[dn]?[au]?', "Tamil Nadu"),
        (r'thami[zl]?[a]?[dn]?[au]?', "Tamil Nadu"),
        (r'tami[zl]?[a]?[dn]?[au]?', "Tamil Nadu"),
        (r'tut[tu]?[uk]?[ku]?[d]?[i]?', "Tamil Nadu"),
        (r'th[uo]?[tu]?[t]?[uh]?[uk]?[ud]?[i]?', "Tamil Nadu"),
        (r'co[iy]mb?[a]?t?[o]?[r]?[e]?', "Tamil Nadu"),
        (r'coy[a]?mb[a]?[t]?[h]?[o]?[r]?[e]?', "Tamil Nadu"),
        (r'tirup[pu]?r', "Tamil Nadu"),
        (r'karaik[ku]?[ud]?[i]?', "Tamil Nadu"),
        (r'ker[ae]l[ae]?', "Kerala"),
        (r'karnat[ai]?k[ae]?', "Karnataka"),
        (r'maharas[ht]{1,2}r[ae]?', "Maharashtra"),
        (r'guj[ar]?[ae]?t', "Gujarat"),
        (r'rajas[ht]{1,2}[ae]?n', "Rajasthan"),
        (r'andhra[ ]?pradesh', "Andhra Pradesh"),
        (r'uttar[ ]?pradesh', "Uttar Pradesh"),
        (r'madhya[ ]?pradesh', "Madhya Pradesh"),
        (r'jammu[ ]?[&]?[ ]?kashmir', "Jammu & Kashmir"),
        (r'j&k', "Jammu & Kashmir"),
        (r'himachal[ ]?pradesh', "Himachal Pradesh"),
        (r'uttarakhand', "Uttarakhand"),
        (r'uttaranchal', "Uttarakhand"),
    ]
    for pattern, state in patterns:
        if re.search(pattern, clean, re.IGNORECASE):
            return state
    return None


def select_address(row):
    for col in ("Permanent Address", "Temporary Address", "Nearby Showroom"):
        val = row.get(col, "")
        if pd.notna(val) and str(val).strip():
            return str(val).strip()
    return ""


@st.cache_data(show_spinner=False)
def find_location(address):
    if pd.isna(address) or str(address).strip() == "":
        return "Unknown", "English"
    original = str(address)
    clean_address = clean_text(original)
    if not clean_address:
        return "Unknown", "English"

    explicit_state_keywords = {
        "kerala": "Kerala", "karnataka": "Karnataka",
        "tamil nadu": "Tamil Nadu", "tamilnadu": "Tamil Nadu",
        "andhra pradesh": "Andhra Pradesh", "andhrapradesh": "Andhra Pradesh",
        "telangana": "Telangana", "maharashtra": "Maharashtra",
        "gujarat": "Gujarat", "west bengal": "West Bengal", "westbengal": "West Bengal",
        "odisha": "Odisha", "orissa": "Odisha",
        "uttar pradesh": "Uttar Pradesh", "uttarpradesh": "Uttar Pradesh",
        "bihar": "Bihar", "rajasthan": "Rajasthan",
        "madhya pradesh": "Madhya Pradesh", "madhyapradesh": "Madhya Pradesh",
        "delhi": "Delhi", "punjab": "Punjab", "haryana": "Haryana",
        "jharkhand": "Jharkhand", "chhattisgarh": "Chhattisgarh",
        "chattisgarh": "Chhattisgarh", "assam": "Assam", "goa": "Goa",
        "puducherry": "Puducherry", "pondicherry": "Puducherry",
        "jammu & kashmir": "Jammu & Kashmir", "jammu and kashmir": "Jammu & Kashmir",
        "himachal pradesh": "Himachal Pradesh", "himachalpradesh": "Himachal Pradesh",
        "uttarakhand": "Uttarakhand", "uttaranchal": "Uttarakhand",
    }
    for keyword, state in explicit_state_keywords.items():
        if keyword in clean_address:
            return state, state_language.get(state, "English")
    for city, state in city_state.items():
        if re.search(r'\b' + re.escape(city) + r'\b', clean_address):
            return state, state_language.get(state, "English")
    variation_state = get_state_from_variation(clean_address)
    if variation_state:
        return variation_state, state_language.get(variation_state, "English")
    fuzzy_state = fuzzy_state_match(clean_address)
    if fuzzy_state:
        return fuzzy_state, state_language.get(fuzzy_state, "English")
    for city, state in city_state.items():
        if city in clean_address:
            return state, state_language.get(state, "English")
    pincode = extract_pincode(original)
    if pincode:
        state = get_state_from_pincode(pincode)
        if state:
            return state, state_language.get(state, "English")
    return "Other", "English"


# -----------------------------------
# Column detection
# -----------------------------------

MOBILE_CANDIDATE_COLS = [
    "mobileno", "mobile no", "mobile no.", "mobilenumber", "mobile number",
    "mobile", "phone", "phone no", "phone no.", "phoneno", "contact",
    "contact no", "contact no.", "contactno", "cell", "cell no",
    "customer mobile", "customer phone", "mob",
]
DISTRICT_CANDIDATE_COLS = [
    "district", "district name", "dist", "dist name", "city", "town",
    "place", "location", "area",
]
LANGUAGE_CANDIDATE_COLS = ["language", "lang", "language name", "lang name"]
INVESTED_CANDIDATE_COLS = [
    "invested amount", "invested", "invested value", "total invested",
    "investment", "investment amount", "amount", "total amount", "net amount",
    "amount paid", "paid amount", "paid", "value", "total value",
    "purchase amount", "sales amount", "grand total", "net value",
    "receipt amount", "scheme amount",
]


def _find_col(columns, candidates):
    lowered = {str(c).strip().lower(): c for c in columns}
    for cand in candidates:
        if cand in lowered:
            return lowered[cand]
    for c in columns:
        cl = str(c).strip().lower()
        for cand in candidates:
            if cand in cl:
                return c
    return None


def find_mobile_col(columns):   return _find_col(columns, MOBILE_CANDIDATE_COLS)
def find_district_col(columns): return _find_col(columns, DISTRICT_CANDIDATE_COLS)
def find_language_col(columns): return _find_col(columns, LANGUAGE_CANDIDATE_COLS)
def find_invested_col(columns): return _find_col(columns, INVESTED_CANDIDATE_COLS)


# -----------------------------------
# 🔧 REFERENCE-BASED PHONE CLEANING
# -----------------------------------

def normalize_mobile(x):
    if x is None or (isinstance(x, float) and pd.isna(x)) or pd.isna(x):
        return None
    s = str(x).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return None

    if "e" in s.lower():
        try:
            s = format(float(s), ".0f")
        except ValueError:
            pass

    s = re.sub(r'\.0+$', '', s)
    s = re.sub(r'[\s\-\(\)\+\.]', '', s)
    s = re.sub(r'[A-Za-z]', '', s)
    s = re.sub(r'\D', '', s)

    if not s:
        return None

    s = re.sub(r'^91([6-9]\d{9})$', r'\1', s)
    s = re.sub(r'^0+(\d{10})$', r'\1', s)

    return s if (len(s) == 10 and s[0] in "6789") else None


def normalize_mobile_loose(x):
    if x is None or (isinstance(x, float) and pd.isna(x)) or pd.isna(x):
        return None
    s = str(x).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return None

    if "e" in s.lower():
        try:
            s = format(float(s), ".0f")
        except ValueError:
            pass
    s = re.sub(r'\.0+$', '', s)
    s = re.sub(r'[\s\-\(\)\+\.]', '', s)
    s = re.sub(r'[A-Za-z]', '', s)
    s = re.sub(r'\D', '', s)

    return s if 10 <= len(s) <= 15 else None


# -----------------------------------
# District helpers
# -----------------------------------

def district_to_state(district):
    if pd.isna(district) or not str(district).strip():
        return None
    clean = clean_text(district)
    for d, st_ in district_state.items():
        if re.search(r'\b' + re.escape(d) + r'\b', clean):
            return st_
    for d, st_ in city_state.items():
        if re.search(r'\b' + re.escape(d) + r'\b', clean):
            return st_
    for d, st_ in district_state.items():
        if d in clean:
            return st_
    for d, st_ in city_state.items():
        if d in clean:
            return st_
    return get_state_from_variation(clean) or fuzzy_state_match(clean)


def infer_district_from_address(addr):
    if pd.isna(addr) or not str(addr).strip():
        return ""
    clean = clean_text(addr)
    for d in district_state.keys():
        if re.search(r'\b' + re.escape(d) + r'\b', clean):
            return d.title()
    for c in city_state.keys():
        if re.search(r'\b' + re.escape(c) + r'\b', clean):
            return c.title()
    return ""


def fill_language_from_district(df_dist, district_col=None, language_col=None):
    if district_col is None:
        district_col = find_district_col(df_dist.columns)
    if language_col is None:
        language_col = find_language_col(df_dist.columns)
    if language_col is None:
        df_dist["Language"] = ""
        language_col = "Language"
    if district_col is None:
        return df_dist
    for idx, row in df_dist.iterrows():
        lang = row.get(language_col, "")
        if pd.notna(lang) and str(lang).strip() and str(lang).strip().lower() != "nan":
            continue
        st_ = district_to_state(row.get(district_col, ""))
        if st_:
            df_dist.at[idx, language_col] = state_language.get(st_, "English")
    return df_dist


# -----------------------------------
# 🚀 CACHED DATA BUILDER
# -----------------------------------

@st.cache_data(show_spinner="Building export sheets (only runs once)...")
def build_all_sheets_cached(df, bss_df, sales_df, mobile_col_f1, invested_col_f1, mobile_set_sales, mobile_set_bss):
    EXPORT_COLS = ["Mobileno", "State", "District", "Language", "Template"]
    
    # ---- Sheet 1: Customer Report rows ----
    if mobile_col_f1:
        mobileno_series = df["_Mobile_Key"].fillna("").astype(str).str.strip()
    else:
        mobileno_series = pd.Series([""] * len(df), index=df.index)

    export_df = pd.DataFrame({
        "Mobileno": mobileno_series,
        "State": df["State"],
        "District": df["_District"].fillna("").astype(str).str.strip(),
        "Language": df["Language"],
        "Template": df["Template"],
    }, index=df.index)

    original_export_len = len(export_df)
    export_df = export_df[export_df["Mobileno"] != ""]
    dropped_count = original_export_len - len(export_df)

    sheet1 = export_df.copy()

    # ---- Sheet 2: Zero Invested from Customer Report ----
    def is_zero_invested(val):
        if pd.isna(val):
            return True
        if isinstance(val, str) and val.strip() == "":
            return True
        try:
            return float(str(val).replace(",", "").strip() or 0) == 0.0
        except ValueError:
            return False

    if invested_col_f1:
        zero_mask = df[invested_col_f1].apply(is_zero_invested)
        valid_mobile_mask = df["_Mobile_Key"].notna()
        sheet2 = export_df[zero_mask.reindex(export_df.index) & valid_mobile_mask.reindex(export_df.index)].copy()
    else:
        zero_mask = ~export_df["Mobileno"].isin(mobile_set_sales.union(mobile_set_bss))
        sheet2 = export_df[zero_mask].copy()

    # ---- Helper: build export rows from a source file ----
    customer_mobile_set = set(df["_Mobile_Key"].dropna().astype(str))

    def build_source_sheets(src_df, label):
        if src_df is None or src_df.empty:
            empty = pd.DataFrame(columns=EXPORT_COLS)
            return empty, empty

        mcol = find_mobile_col(src_df.columns)
        dcol = find_district_col(src_df.columns)
        lcol = find_language_col(src_df.columns)

        if not mcol:
            empty = pd.DataFrame(columns=EXPORT_COLS)
            return empty, src_df.assign(
                Mobileno="", State="", District="", Language="", Template=""
            )[EXPORT_COLS].reset_index(drop=True)

        src = fill_language_from_district(src_df.copy(), district_col=dcol, language_col=lcol)

        rows = []
        for _, row in src.iterrows():
            raw_mobile = row.get(mcol, "")
            key = normalize_mobile(raw_mobile)
            dist_val = str(row.get(dcol, "") or "").strip() if dcol else ""
            lang_val = str(row.get(lcol, "") or "").strip() if lcol else ""
            state_val = district_to_state(dist_val) or ""
            if not lang_val:
                lang_val = state_language.get(state_val, "English") if state_val else "English"
            template_val = "Tamil" if lang_val == "Tamil" else "English"
            rows.append({
                "Mobileno": key or "",
                "State": state_val,
                "District": dist_val,
                "Language": lang_val,
                "Template": template_val,
                "_key": key,
            })

        out = pd.DataFrame(rows)
        out = out[out["Mobileno"] != ""]

        matched_mask = out["_key"].apply(lambda k: bool(k) and k in customer_mobile_set)
        matched = out[matched_mask][EXPORT_COLS].reset_index(drop=True)
        unmatched = out[~matched_mask][EXPORT_COLS].reset_index(drop=True)
        return matched, unmatched

    sheet3_sales_matched, sheet5_sales_unmatched = build_source_sheets(sales_df, "Sales Purchase")
    sheet4_bss_matched, sheet6_bss_unmatched = build_source_sheets(bss_df, "Bss Receipt")

    # Determine invested mobiles
    if invested_col_f1:
        invested_mask_cust = ~df[invested_col_f1].apply(is_zero_invested)
        invested_mobiles = set(df.loc[invested_mask_cust, "_Mobile_Key"].dropna().astype(str))
    else:
        invested_mobiles = mobile_set_sales.union(mobile_set_bss)

    sheet7_purch_invested = sheet3_sales_matched[
        sheet3_sales_matched["Mobileno"].isin(invested_mobiles)
    ].reset_index(drop=True)
    
    sheet8_purch_no_invested = sheet3_sales_matched[
        ~sheet3_sales_matched["Mobileno"].isin(invested_mobiles)
    ].reset_index(drop=True)

    sheet9_bss_invested = sheet4_bss_matched[
        sheet4_bss_matched["Mobileno"].isin(invested_mobiles)
    ].reset_index(drop=True)
    
    sheet10_bss_no_invested = sheet4_bss_matched[
        ~sheet4_bss_matched["Mobileno"].isin(invested_mobiles)
    ].reset_index(drop=True)

    return {
        SHEET_META[0][1]: sheet1,
        SHEET_META[1][1]: sheet2,
        SHEET_META[2][1]: sheet3_sales_matched,
        SHEET_META[3][1]: sheet4_bss_matched,
        SHEET_META[4][1]: sheet5_sales_unmatched,
        SHEET_META[5][1]: sheet6_bss_unmatched,
        SHEET_META[6][1]: sheet7_purch_invested,
        SHEET_META[7][1]: sheet8_purch_no_invested,
        SHEET_META[8][1]: sheet9_bss_invested,
        SHEET_META[9][1]: sheet10_bss_no_invested,
    }, dropped_count


# -----------------------------------
# Sidebar
# -----------------------------------

with st.sidebar:
    st.header("ℹ️ How it works")
    st.markdown(
        "1. Upload the **customer report** (required) — `.xlsx` / `.csv`\n"
        "2. Optionally upload **Bss Receipt** and **Sales Purchase** files\n"
        "3. State / District / Language detected from address\n"
        "4. **Unknown** customers re-checked **sequentially**:\n"
        "   - First: Bss Receipt (`Mobileno → District → Language`)\n"
        "   - Then: Sales Purchase\n"
        "5. Download **20 Excel files** (10 reports × 2 languages)\n"
        "   - `Tamil` customers go to the Tamil file\n"
        "   - All other languages go to the English file\n"
        "   - **Or download all 20 in one ZIP**"
    )
    st.divider()
    st.subheader("Customer Report")
    st.markdown("- **Permanent Address** *(required)*\n- Temporary Address\n"
                "- Nearby Showroom\n- Mobile column (any common name) — enables fallback lookup\n"
                "- Invested column — enables Sheet 2, 7, 8, 9, 10 (Invested splitting)")
    st.subheader("Bss Receipt / Sales Purchase")
    st.markdown("- `Sr` `Date` `Mobileno` `District` `Language`\n"
                "- District → State → Language auto-filled if Language is blank\n"
                "- Unmatched sheets pull District/State/Language from their own file row")
    st.divider()
    st.caption("📱 Phone cleaning follows reference logic: strips spaces/hyphens/brackets/plus, "
               "removes letters, +91 → 10-digit (only if valid mobile), leading zeros stripped.")


# -----------------------------------
# Uploaders
# -----------------------------------

st.subheader("1) Upload Files")
col_u1, col_u2, col_u3 = st.columns(3)
with col_u1:
    uploaded_file = st.file_uploader("Customer Report *", type=["xlsx", "csv"], key="f1")
with col_u2:
    uploaded_bss = st.file_uploader("Bss Receipt (optional)", type=["xlsx", "csv"], key="f2")
with col_u3:
    uploaded_sales = st.file_uploader("Sales Purchase (optional)", type=["xlsx", "csv"], key="f3")

if not uploaded_file:
    st.info("👋 Upload the Customer Report to get started. Bss Receipt and Sales Purchase are optional.")
    st.stop()


def load_any(f):
    if f is None:
        return None
    try:
        if f.name.lower().endswith(".xlsx"):
            return pd.read_excel(f, dtype=str)
        return pd.read_csv(f, dtype=str)
    except Exception as e:
        st.error(f"❌ Couldn't read `{f.name}`: {e}")
        return None


# ---- Load customer report ----
df = load_any(uploaded_file)
if df is None or df.empty:
    st.stop()
if "Permanent Address" not in df.columns:
    st.error("❌ 'Permanent Address' column missing in Customer Report. Found: " + ", ".join(df.columns))
    st.stop()

for c in ("Temporary Address", "Nearby Showroom"):
    if c not in df.columns:
        df[c] = ""

df["Final Address"] = df.apply(select_address, axis=1)


# ---- Load Bss + Sales ----
bss_df = load_any(uploaded_bss) if uploaded_bss else None
sales_df = load_any(uploaded_sales) if uploaded_sales else None

mobile_col_f1 = find_mobile_col(df.columns)
invested_col_f1 = find_invested_col(df.columns)

if mobile_col_f1:
    df["_Mobile_Key"] = df[mobile_col_f1].apply(normalize_mobile)
else:
    df["_Mobile_Key"] = None


# -----------------------------------
# Address-based detection on Customer Report
# -----------------------------------

progress_bar = st.progress(0, text="Detecting states and languages...")
states, languages = [], []
total_rows = len(df)
chunk = max(1, total_rows // 100)
for i, addr in enumerate(df["Final Address"]):
    state, language = find_location(addr)
    states.append(state)
    languages.append(language)
    if i % chunk == 0 or i == total_rows - 1:
        progress_bar.progress((i + 1) / total_rows,
                              text=f"Detecting states and languages... ({i+1:,}/{total_rows:,})")
progress_bar.empty()

df["State"] = states
df["Language"] = languages


# -----------------------------------
# Build lookups for fallback resolution
# -----------------------------------

def build_mobile_lookup(src_df, src_name):
    stats = {"file": src_name, "rows": 0, "mobile_col": None,
             "district_col": None, "language_col": None,
             "unique_mobiles": 0, "with_state": 0}
    if src_df is None or src_df.empty:
        return {}, stats

    stats["rows"] = len(src_df)
    mcol = find_mobile_col(src_df.columns)
    dcol = find_district_col(src_df.columns)
    lcol = find_language_col(src_df.columns)
    stats["mobile_col"] = mcol
    stats["district_col"] = dcol
    stats["language_col"] = lcol

    if not mcol:
        return {}, stats

    src = fill_language_from_district(src_df.copy(), district_col=dcol, language_col=lcol)

    lookup = {}
    for _, row in src.iterrows():
        key = normalize_mobile(row.get(mcol))
        if not key:
            continue
        dist = row.get(dcol, "") if dcol else ""
        lang = row.get(lcol, "") if lcol else ""
        state_ = district_to_state(dist)
        if not (pd.notna(lang) and str(lang).strip()):
            lang = state_language.get(state_, "English") if state_ else "English"
        lookup.setdefault(key, {
            "District": str(dist).strip() if pd.notna(dist) else "",
            "State": state_,
            "Language": lang,
            "Source": src_name,
        })
    stats["unique_mobiles"] = len(lookup)
    stats["with_state"] = sum(1 for v in lookup.values() if v.get("State"))
    return lookup, stats


def build_mobile_set(src_df):
    if src_df is None or src_df.empty:
        return set()
    mcol = find_mobile_col(src_df.columns)
    if not mcol:
        return set()
    out = set()
    for _, row in src_df.iterrows():
        k = normalize_mobile(row.get(mcol))
        if k:
            out.add(k)
    return out


def _apply_lookup_to_unknown(df, lookup, source_label):
    resolved = 0
    for idx, row in df.iterrows():
        if row["State"] not in ("Unknown", "Other"):
            continue
        key = row.get("_Mobile_Key")
        if not key:
            continue
        hit = lookup.get(key)
        if not hit:
            continue
        df.at[idx, "State"] = hit.get("State") or row["State"]
        df.at[idx, "Language"] = hit.get("Language") or row["Language"]
        if hit.get("District") and str(hit["District"]).strip():
            df.at[idx, "_District"] = str(hit["District"]).strip()
        df.at[idx, "_MatchedFrom"] = source_label
        resolved += 1
    return resolved


lookup_bss, stats_bss = build_mobile_lookup(bss_df, "Bss Receipt")
lookup_sales, stats_sales = build_mobile_lookup(sales_df, "Sales Purchase")

mobile_set_bss = build_mobile_set(bss_df)
mobile_set_sales = build_mobile_set(sales_df)

if not mobile_col_f1:
    st.warning("⚠️ No mobile column found in the Customer Report — mobile fallback skipped.")
else:
    valid_mobiles = df['_Mobile_Key'].notna().sum()
    st.caption(f"📱 Mobile column in Customer Report: `{mobile_col_f1}` — "
               f"{valid_mobiles:,} valid 10-digit mobiles after normalization")

with st.expander("🔎 Mobile Normalization Check (verify +91 → 10 digit)", expanded=False):
    if mobile_col_f1:
        sample = pd.DataFrame({
            "Raw value": df[mobile_col_f1].head(10).astype(str).values,
            "Normalized": df["_Mobile_Key"].head(10).values,
        })
        st.dataframe(sample, hide_index=True, width="stretch")
    else:
        st.info("No mobile column detected in Customer Report.")

df["_MatchedFrom"] = ""
df["_District"] = ""

resolved_bss = _apply_lookup_to_unknown(df, lookup_bss, "Bss Receipt")
resolved_sales = _apply_lookup_to_unknown(df, lookup_sales, "Sales Purchase")
total_resolved = resolved_bss + resolved_sales

df["Template"] = df["Language"].apply(lambda lang: "Tamil" if lang == "Tamil" else "English")

blank_district_mask = df["_District"].fillna("").astype(str).str.strip() == ""
df.loc[blank_district_mask, "_District"] = (
    df.loc[blank_district_mask, "Final Address"].apply(infer_district_from_address)
)

st.success(f"✅ State, District, Language & Template added "
           f"(Bss resolved {resolved_bss:,} • Sales resolved {resolved_sales:,} • "
           f"total {total_resolved:,} unknowns filled)")


# -----------------------------------
# Diagnostics
# -----------------------------------

with st.expander("🔍 Match Diagnostics (mobile → district lookup)", expanded=False):
    st.markdown("**Lookups built**")
    st.dataframe(pd.DataFrame([
        {"File": stats_bss["file"], "Rows": stats_bss["rows"],
         "Mobile col": stats_bss["mobile_col"] or "— not found —",
         "District col": stats_bss["district_col"] or "—",
         "Language col": stats_bss["language_col"] or "—",
         "Unique mobiles": stats_bss["unique_mobiles"],
         "Mobiles w/ State": stats_bss["with_state"]},
        {"File": stats_sales["file"], "Rows": stats_sales["rows"],
         "Mobile col": stats_sales["mobile_col"] or "— not found —",
         "District col": stats_sales["district_col"] or "—",
         "Language col": stats_sales["language_col"] or "—",
         "Unique mobiles": stats_sales["unique_mobiles"],
         "Mobiles w/ State": stats_sales["with_state"]},
    ]), width="stretch", hide_index=True)

    st.markdown("**Sequential resolution of Unknown rows**")
    st.dataframe(pd.DataFrame([
        {"Step": "1", "Source": "Bss Receipt",    "Resolved": resolved_bss},
        {"Step": "2", "Source": "Sales Purchase", "Resolved": resolved_sales},
        {"Step": "—", "Source": "TOTAL resolved", "Resolved": total_resolved},
    ]), width="stretch", hide_index=True)

    still_unknown = (df["State"] == "Unknown").sum()
    st.caption(f"Still Unknown after both steps: **{still_unknown:,}** "
               f"({still_unknown/len(df)*100:.1f}% of {len(df):,})")


# -----------------------------------
# Metrics
# -----------------------------------

total_customers = len(df)
tamil_customers = (df["Language"] == "Tamil").sum()
unknown = (df["State"] == "Unknown").sum()
other_languages = (df["Language"] != "Tamil").sum()
non_tamil_identified = other_languages - unknown
detected = total_customers - unknown
detection_rate = (detected / total_customers * 100) if total_customers else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("📊 Total Customers", f"{total_customers:,}")
c2.metric("🇮🇳 Tamil", f"{tamil_customers:,}", delta=f"{tamil_customers/total_customers*100:.1f}%")
c3.metric("🌍 Other Languages", f"{non_tamil_identified:,}")
c4.metric("❓ Unknown", f"{unknown:,}", delta=f"{unknown/total_customers*100:.1f}%", delta_color="inverse")

st.progress(detection_rate / 100 if detection_rate > 0 else 0)
if detection_rate < 50:
    st.error(f"⚠️ **Low Detection Rate: {detection_rate:.1f}%** — Only {detected:,} out of {total_customers:,} identified")
else:
    st.success(f"✅ **Detection Rate: {detection_rate:.1f}%** — {detected:,} customers identified")

st.divider()


# -----------------------------------
# Tabs
# -----------------------------------

tab_overview, tab_results, tab_unknown, tab_download = st.tabs(
    ["📊 Overview", "📋 Full Results", "⚠️ Unknown Addresses", "📥 Final Export"]
)

with tab_overview:
    st.subheader("Customer Breakdown")
    st.dataframe(pd.DataFrame({
        "Category": ["Tamil Customers", "Other Languages (Identified)", "Unknown", "Total"],
        "Count": [tamil_customers, non_tamil_identified, unknown, total_customers],
        "Percentage": [f"{v/total_customers*100:.1f}%" for v in
                       [tamil_customers, non_tamil_identified, unknown, total_customers]],
    }), width="stretch", hide_index=True)

    ca, cb = st.columns(2)
    with ca:
        st.markdown("**Language Distribution**")
        st.bar_chart(df["Language"].value_counts())
    with cb:
        st.markdown("**Top 10 States**")
        st.bar_chart(df["State"].value_counts().head(10))

with tab_results:
    st.subheader("Customer Result")
    fc1, fc2, fc3 = st.columns([2, 1, 1])
    with fc1:
        search_term = st.text_input("🔎 Search address", placeholder="Type part of an address...")
    with fc2:
        state_filter = st.selectbox("Filter by state", ["All"] + sorted(df["State"].unique().tolist()))
    with fc3:
        language_filter = st.selectbox("Filter by language", ["All"] + sorted(df["Language"].unique().tolist()))

    fdf = df.copy()
    if search_term:
        fdf = fdf[fdf["Final Address"].str.contains(search_term, case=False, na=False)]
    if state_filter != "All":
        fdf = fdf[fdf["State"] == state_filter]
    if language_filter != "All":
        fdf = fdf[fdf["Language"] == language_filter]

    display_cols = [c for c in ["_Mobile_Key", "Permanent Address", "Temporary Address",
                                "Nearby Showroom", "Final Address", "State", "_District",
                                "Language", "Template", "_MatchedFrom"] if c in fdf.columns]
    st.caption(f"Showing {len(fdf):,} of {total_customers:,} customers")
    st.dataframe(
        fdf[display_cols].rename(columns={"_Mobile_Key": "Mobileno",
                                          "_District": "District",
                                          "_MatchedFrom": "Matched From"}),
        width="stretch", hide_index=True
    )

with tab_unknown:
    unknown_df = df[df["State"] == "Unknown"]
    if len(unknown_df) > 0:
        st.warning(f"⚠️ **{len(unknown_df):,} customers with unknown state** "
                   f"({len(unknown_df)/total_customers*100:.1f}% of total)")
        unique_unknown = unknown_df["Final Address"].value_counts()
        st.write("**Top 20 unknown addresses:**")
        st.dataframe(pd.DataFrame({"Address": unique_unknown.index[:20],
                                   "Count": unique_unknown.values[:20]}),
                     width="stretch", hide_index=True)
        st.download_button("📥 Download Unknown Addresses (CSV)",
                           unknown_df[["Final Address"]].to_csv(index=False).encode("utf-8"),
                           "unknown_addresses.csv", "text/csv")
    else:
        st.success("🎉 No unknown addresses found! 100% detection rate!")


# -----------------------------------
# Final Export: 10 reports × 2 languages = 20 individual Excel files
# -----------------------------------

with tab_download:
    st.subheader("📥 Final Export — 20 Individual Excel Files")
    st.caption("Each of the 10 reports is split into a **Tamil** file and an **English** file "
               "(Tamil customers → Tamil file; all other languages → English file). "
               "Download all files at once in a ZIP, or download individually below.")

    # 🚀 Build all 10 base reports once
    all_sheets_dict, dropped_count = build_all_sheets_cached(
        df, bss_df, sales_df, mobile_col_f1, invested_col_f1, mobile_set_sales, mobile_set_bss
    )

    if dropped_count > 0:
        st.info(f"🧹 **Clean-up notice:** {dropped_count:,} rows with blank/empty phone numbers "
                f"were automatically removed from the final report.")
    
    if invested_col_f1:
        st.caption(f"Sheets 2, 7, 8, 9, 10 use Customer Report column **`{invested_col_f1}`** — rows where value is `0` or blank are treated as *No Invested*.")
    else:
        st.caption("No invested column found in Customer Report — fallback used: *Invested* = mobiles present in Sales or Bss; *No Invested* = not present in either.")

    st.info("ℹ️ **Sheets 3–6 are rows from the Sales/Bss files themselves.** "
            "District/State/Language on those rows come from that row's own data — "
            "so unmatched rows still display their own District.")

    st.divider()

    # ---- Summary metrics for the 20 files ----
    st.markdown("### 📊 Export Summary (20 Files)")
    summary_rows = []
    for _, full_title in SHEET_META:
        data = all_sheets_dict[full_title]
        tamil_data = data[data["Language"] == "Tamil"]
        english_data = data[data["Language"] != "Tamil"]
        summary_rows.append({
            "Report": full_title,
            "Tamil Rows": len(tamil_data),
            "English Rows": len(english_data),
            "Total Rows": len(data),
        })
    st.dataframe(pd.DataFrame(summary_rows), width="stretch", hide_index=True)

    st.divider()

    # ---- Download ALL 20 files as a single ZIP ----
    st.markdown("### 📦 Download All 20 Files in One ZIP")
    st.caption("The ZIP contains every report split by language: "
               "`<Report> - Tamil.xlsx` and `<Report> - English.xlsx`.")

    zip_all_out = BytesIO()
    total_files_in_zip = 0
    with zipfile.ZipFile(zip_all_out, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for _, full_title in SHEET_META:
            data = all_sheets_dict[full_title]
            for lang_label, mask in [("Tamil", data["Language"] == "Tamil"),
                                     ("English", data["Language"] != "Tamil")]:
                subset = data[mask].reset_index(drop=True)
                if len(subset) == 0:
                    continue
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    subset.to_excel(writer, sheet_name="Sheet1", index=False, startrow=0)
                zip_file.writestr(f"{full_title} - {lang_label}.xlsx", buf.getvalue())
                total_files_in_zip += 1

    st.download_button(
        label=f"📦 Download All {total_files_in_zip} Files (.zip)",
        data=zip_all_out.getvalue(),
        file_name="All_Customer_Reports_Tamil_And_English.zip",
        mime="application/zip",
        key="dl_all_zip",
        width="stretch",
    )
    st.caption(f"ZIP contains **{total_files_in_zip}** Excel files "
               f"(empty language subsets are skipped).")

    st.divider()

    # ---- 20 individual download tabs ----
    st.markdown("### 📥 Download Individual Files")
    
    tabs = st.tabs([meta[1] for meta in SHEET_META])

    for tab, (_, full_title) in zip(tabs, SHEET_META):
        data = all_sheets_dict[full_title]
        tamil_data = data[data["Language"] == "Tamil"].reset_index(drop=True)
        english_data = data[data["Language"] != "Tamil"].reset_index(drop=True)

        with tab:
            st.markdown(f"**{full_title}** — {len(data):,} total rows "
                        f"({len(tamil_data):,} Tamil • {len(english_data):,} English)")

            st.markdown(f"#### 🇮🇳 `{full_title} - Tamil.xlsx`")
            if len(tamil_data) > 0:
                tamil_buf = BytesIO()
                with pd.ExcelWriter(tamil_buf, engine="openpyxl") as writer:
                    tamil_data.to_excel(writer, sheet_name="Sheet1", index=False, startrow=0)
                st.download_button(
                    label=f"📥 Download {full_title} - Tamil.xlsx ({len(tamil_data):,} rows)",
                    data=tamil_buf.getvalue(),
                    file_name=f"{full_title} - Tamil.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_tamil_{full_title}",
                    width="stretch",
                )
                with st.expander("Preview Tamil", expanded=False):
                    st.dataframe(tamil_data.head(20), hide_index=True, width="stretch")
            else:
                st.info("No Tamil rows in this report.")
                st.button(
                    f"📥 Download {full_title} - Tamil.xlsx (0 rows)",
                    disabled=True,
                    key=f"dl_tamil_empty_{full_title}",
                    width="stretch",
                )

            st.markdown(f"#### 🌍 `{full_title} - English.xlsx`")
            if len(english_data) > 0:
                english_buf = BytesIO()
                with pd.ExcelWriter(english_buf, engine="openpyxl") as writer:
                    english_data.to_excel(writer, sheet_name="Sheet1", index=False, startrow=0)
                st.download_button(
                    label=f"📥 Download {full_title} - English.xlsx ({len(english_data):,} rows)",
                    data=english_buf.getvalue(),
                    file_name=f"{full_title} - English.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_english_{full_title}",
                    width="stretch",
                )
                with st.expander("Preview English", expanded=False):
                    st.dataframe(english_data.head(20), hide_index=True, width="stretch")
            else:
                st.info("No English rows in this report.")
                st.button(
                    f"📥 Download {full_title} - English.xlsx (0 rows)",
                    disabled=True,
                    key=f"dl_english_empty_{full_title}",
                    width="stretch",
                )