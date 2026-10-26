import streamlit as st
import pandas as pd
from io import BytesIO
import re
import zipfile
import hashlib
from functools import lru_cache

# -----------------------------------
# Streamlit Configuration
# -----------------------------------

st.set_page_config(
    page_title="Customer Location Finder",
    page_icon="📍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------
# 🎨 Look & feel (CSS only – no logic)
# -----------------------------------

st.markdown(
    """
<style>
.block-container {padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1400px;}

/* Hero header */
.hero {
    background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 55%, #db2777 100%);
    padding: 1.7rem 2rem; border-radius: 20px; color: #fff;
    margin-bottom: 1.4rem; box-shadow: 0 12px 32px rgba(79,70,229,.28);
}
.hero h1 {margin: 0 0 .35rem 0; font-size: 2.1rem; font-weight: 800; letter-spacing: -.5px; color:#fff;}
.hero p {margin: 0; font-size: .98rem; opacity: .93; line-height: 1.55; color:#fff;}
.hero .pill {
    display: inline-block; background: rgba(255,255,255,.18); border-radius: 999px;
    padding: .18rem .7rem; font-size: .78rem; margin: .7rem .4rem 0 0; font-weight: 600;
}

/* Section titles */
.section-title {font-size: 1.15rem; font-weight: 700; margin: .2rem 0 .6rem 0;}
.step-badge {
    display:inline-block; background:#4f46e5; color:#fff; border-radius:50%;
    width:26px; height:26px; line-height:26px; text-align:center; font-size:.85rem;
    font-weight:700; margin-right:.5rem;
}

/* Metric cards */
div[data-testid="stMetric"] {
    background: rgba(128,128,128,.07); border: 1px solid rgba(128,128,128,.20);
    padding: 16px 20px; border-radius: 16px; transition: transform .15s ease, box-shadow .15s ease;
}
div[data-testid="stMetric"]:hover {transform: translateY(-2px); box-shadow: 0 8px 20px rgba(0,0,0,.10);}
div[data-testid="stMetricValue"] {font-weight: 800;}

/* Buttons */
.stDownloadButton button, .stButton button {
    border-radius: 12px; font-weight: 600; padding: .55rem 1rem;
    border: 1px solid rgba(128,128,128,.3); transition: all .15s ease;
}
.stDownloadButton button:hover {
    border-color: #4f46e5; color: #4f46e5; transform: translateY(-1px);
    box-shadow: 0 6px 16px rgba(79,70,229,.18);
}

/* Tabs */
button[data-baseweb="tab"] {font-weight: 600; border-radius: 10px 10px 0 0;}

/* Uploaders & containers */
section[data-testid="stFileUploaderDropzone"] {border-radius: 14px; border: 2px dashed rgba(124,58,237,.45);}
div[data-testid="stExpander"] {border-radius: 14px; border: 1px solid rgba(128,128,128,.22);}
div[data-testid="stDataFrame"] {border-radius: 12px; overflow: hidden;}

/* Sidebar */
section[data-testid="stSidebar"] {border-right: 1px solid rgba(128,128,128,.2);}
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="hero">
  <h1>📍 Customer State &amp; Language Finder</h1>
  <p>Upload a customer report plus optional <b>Bss Receipt</b> / <b>Sales Purchase</b> files.
  Unknown customers are resolved sequentially: Bss Receipt → Sales Purchase.
  Export includes 10 sheets with Mobileno, State, District, Language, Template.</p>
  <span class="pill">🧹 Blank phone rows auto-removed</span>
  <span class="pill">⚡ Cached &amp; fast</span>
  <span class="pill">📦 20 Excel files + ZIP</span>
</div>
""",
    unsafe_allow_html=True,
)


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
# ⚡ Pre-compiled regex (same patterns, same order → same results, just faster)
# -----------------------------------

_CITY_WORD_RE = [(re.compile(r'\b' + re.escape(c) + r'\b'), s) for c, s in city_state.items()]
_DISTRICT_WORD_RE = [(re.compile(r'\b' + re.escape(d) + r'\b'), s) for d, s in district_state.items()]

_FUZZY_PATTERNS = [
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
_FUZZY_COMPILED = [(re.compile(p, re.IGNORECASE), s) for p, s in _FUZZY_PATTERNS]

_EXPLICIT_STATE_KEYWORDS = {
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
    for pattern, state in _FUZZY_COMPILED:
        if pattern.search(clean):
            return state
    return None


def find_location_uncached(address):
    if pd.isna(address) or str(address).strip() == "":
        return "Unknown", "English"
    original = str(address)
    clean_address = clean_text(original)
    if not clean_address:
        return "Unknown", "English"

    for keyword, state in _EXPLICIT_STATE_KEYWORDS.items():
        if keyword in clean_address:
            return state, state_language.get(state, "English")
    for pattern, state in _CITY_WORD_RE:
        if pattern.search(clean_address):
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


find_location = lru_cache(maxsize=None)(find_location_uncached)


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


_normalize_mobile_cached = lru_cache(maxsize=None)(normalize_mobile)


# -----------------------------------
# District helpers
# -----------------------------------

def district_to_state_uncached(district):
    if pd.isna(district) or not str(district).strip():
        return None
    clean = clean_text(district)
    for pattern, st_ in _DISTRICT_WORD_RE:
        if pattern.search(clean):
            return st_
    for pattern, st_ in _CITY_WORD_RE:
        if pattern.search(clean):
            return st_
    for d, st_ in district_state.items():
        if d in clean:
            return st_
    for d, st_ in city_state.items():
        if d in clean:
            return st_
    return get_state_from_variation(clean) or fuzzy_state_match(clean)


district_to_state = lru_cache(maxsize=None)(district_to_state_uncached)


@lru_cache(maxsize=None)
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

    df_dist[language_col] = df_dist[language_col].astype(object)
    lang = df_dist[language_col]
    lang_str = lang.astype(str).str.strip()
    # rows whose language is blank / "nan" get filled from district (same rule as before)
    blank = lang.isna() | (lang_str == "") | (lang_str.str.lower() == "nan")
    if not blank.any():
        return df_dist

    def _lang_for(d):
        st_ = district_to_state(d)
        return state_language.get(st_, "English") if st_ else None

    new_vals = df_dist.loc[blank, district_col].map(_lang_for)
    valid = new_vals.notna()
    if valid.any():
        df_dist.loc[new_vals.index[valid], language_col] = new_vals[valid]
    return df_dist


def _blank_safe_str(v):
    return str(v or "").strip()


# -----------------------------------
# 🚀 EXPORT BUILDER (same logic, vectorized)
# -----------------------------------

def build_all_sheets(df, bss_df, sales_df, mobile_col_f1, invested_col_f1, mobile_set_sales, mobile_set_bss):
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

    zero_series = df[invested_col_f1].apply(is_zero_invested) if invested_col_f1 else None

    if invested_col_f1:
        zero_mask = zero_series
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

        keys = src[mcol].map(_normalize_mobile_cached)
        if dcol:
            dist_val = src[dcol].map(_blank_safe_str)
        else:
            dist_val = pd.Series([""] * len(src), index=src.index)
        if lcol:
            lang_val = src[lcol].map(_blank_safe_str)
        else:
            lang_val = pd.Series([""] * len(src), index=src.index)

        state_val = dist_val.map(lambda d: district_to_state(d) or "")
        fallback = state_val.map(lambda s: state_language.get(s, "English") if s else "English")
        lang_val = lang_val.where(lang_val != "", fallback)
        template_val = lang_val.map(lambda l: "Tamil" if l == "Tamil" else "English")

        out = pd.DataFrame({
            "Mobileno": keys.fillna(""),
            "State": state_val,
            "District": dist_val,
            "Language": lang_val,
            "Template": template_val,
            "_key": keys,
        })
        out = out[out["Mobileno"] != ""]

        matched_mask = out["_key"].isin(customer_mobile_set)
        matched = out[matched_mask][EXPORT_COLS].reset_index(drop=True)
        unmatched = out[~matched_mask][EXPORT_COLS].reset_index(drop=True)
        return matched, unmatched

    sheet3_sales_matched, sheet5_sales_unmatched = build_source_sheets(sales_df, "Sales Purchase")
    sheet4_bss_matched, sheet6_bss_unmatched = build_source_sheets(bss_df, "Bss Receipt")

    # Determine invested mobiles
    if invested_col_f1:
        invested_mask_cust = ~zero_series
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
# Lookups for fallback resolution
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

    keys = src[mcol].map(_normalize_mobile_cached).tolist()
    dists = src[dcol].tolist() if dcol else [""] * len(src)
    langs = src[lcol].tolist() if lcol else [""] * len(src)

    lookup = {}
    for key, dist, lang in zip(keys, dists, langs):
        if not key or key in lookup:
            continue
        state_ = district_to_state(dist)
        if not (pd.notna(lang) and str(lang).strip()):
            lang = state_language.get(state_, "English") if state_ else "English"
        lookup[key] = {
            "District": str(dist).strip() if pd.notna(dist) else "",
            "State": state_,
            "Language": lang,
            "Source": src_name,
        }
    stats["unique_mobiles"] = len(lookup)
    stats["with_state"] = sum(1 for v in lookup.values() if v.get("State"))
    return lookup, stats


def build_mobile_set(src_df):
    if src_df is None or src_df.empty:
        return set()
    mcol = find_mobile_col(src_df.columns)
    if not mcol:
        return set()
    return {k for k in src_df[mcol].map(_normalize_mobile_cached) if k}


def _apply_lookup_to_unknown(df, lookup, source_label):
    if not lookup:
        return 0
    cand = df["State"].isin(("Unknown", "Other")) & df["_Mobile_Key"].isin(list(lookup.keys()))
    idxs = df.index[cand]
    resolved = 0
    for idx, key in zip(idxs, df.loc[idxs, "_Mobile_Key"]):
        hit = lookup.get(key)
        if not hit:
            continue
        df.at[idx, "State"] = hit.get("State") or df.at[idx, "State"]
        df.at[idx, "Language"] = hit.get("Language") or df.at[idx, "Language"]
        if hit.get("District") and str(hit["District"]).strip():
            df.at[idx, "_District"] = str(hit["District"]).strip()
        df.at[idx, "_MatchedFrom"] = source_label
        resolved += 1
    return resolved


# -----------------------------------
# File reading
# -----------------------------------

def _read_df(name, data):
    if name.lower().endswith(".xlsx"):
        return pd.read_excel(BytesIO(data), dtype=str)
    return pd.read_csv(BytesIO(data), dtype=str)


# -----------------------------------
# 🚀 FULL PIPELINE — runs ONCE per upload, then served from cache
# (filters / tabs / searches no longer re-run any heavy work)
# -----------------------------------

@st.cache_resource(show_spinner="⚙️ Detecting states, languages & resolving unknowns (runs once per upload)...")
def run_pipeline(token, _cust_bytes, cust_name, _bss_bytes, bss_name, _sales_bytes, sales_name):
    res = {"fatal": None, "warnings": []}

    # ---- Load customer report ----
    try:
        df = _read_df(cust_name, _cust_bytes)
    except Exception as e:
        res["fatal"] = f"❌ Couldn't read `{cust_name}`: {e}"
        return res
    if df is None or df.empty:
        res["fatal"] = "❌ The Customer Report is empty."
        return res
    if "Permanent Address" not in df.columns:
        res["fatal"] = "❌ 'Permanent Address' column missing in Customer Report. Found: " + ", ".join(map(str, df.columns))
        return res

    for c in ("Temporary Address", "Nearby Showroom"):
        if c not in df.columns:
            df[c] = ""

    # Final Address (same rule as select_address, vectorized)
    final_addr = pd.Series("", index=df.index, dtype=object)
    for col in reversed(("Permanent Address", "Temporary Address", "Nearby Showroom")):
        s = df[col].astype(str).str.strip()
        valid = df[col].notna() & (s != "")
        final_addr = final_addr.where(~valid, s)
    df["Final Address"] = final_addr

    # ---- Load Bss + Sales ----
    def _safe(name, data):
        if data is None:
            return None
        try:
            return _read_df(name, data)
        except Exception as e:
            res["warnings"].append(f"❌ Couldn't read `{name}`: {e}")
            return None

    bss_df = _safe(bss_name, _bss_bytes)
    sales_df = _safe(sales_name, _sales_bytes)

    mobile_col_f1 = find_mobile_col(df.columns)
    invested_col_f1 = find_invested_col(df.columns)

    if mobile_col_f1:
        df["_Mobile_Key"] = df[mobile_col_f1].map(_normalize_mobile_cached)
    else:
        df["_Mobile_Key"] = None

    # ---- Address-based detection (unique addresses only) ----
    uniq = df["Final Address"].unique()
    loc_map = {a: find_location(a) for a in uniq}
    df["State"] = df["Final Address"].map(lambda a: loc_map[a][0])
    df["Language"] = df["Final Address"].map(lambda a: loc_map[a][1])

    # ---- Lookups ----
    lookup_bss, stats_bss = build_mobile_lookup(bss_df, "Bss Receipt")
    lookup_sales, stats_sales = build_mobile_lookup(sales_df, "Sales Purchase")

    mobile_set_bss = build_mobile_set(bss_df)
    mobile_set_sales = build_mobile_set(sales_df)

    valid_mobiles = int(df["_Mobile_Key"].notna().sum()) if mobile_col_f1 else 0
    sample = None
    if mobile_col_f1:
        sample = pd.DataFrame({
            "Raw value": df[mobile_col_f1].head(10).astype(str).values,
            "Normalized": df["_Mobile_Key"].head(10).values,
        })

    df["_MatchedFrom"] = ""
    df["_District"] = ""

    resolved_bss = _apply_lookup_to_unknown(df, lookup_bss, "Bss Receipt")
    resolved_sales = _apply_lookup_to_unknown(df, lookup_sales, "Sales Purchase")

    df["Template"] = df["Language"].map(lambda lang: "Tamil" if lang == "Tamil" else "English")

    blank_district_mask = df["_District"].fillna("").astype(str).str.strip() == ""
    df.loc[blank_district_mask, "_District"] = (
        df.loc[blank_district_mask, "Final Address"].map(infer_district_from_address)
    )

    res.update({
        "df": df, "bss_df": bss_df, "sales_df": sales_df,
        "mobile_col_f1": mobile_col_f1, "invested_col_f1": invested_col_f1,
        "mobile_set_bss": mobile_set_bss, "mobile_set_sales": mobile_set_sales,
        "stats_bss": stats_bss, "stats_sales": stats_sales,
        "resolved_bss": resolved_bss, "resolved_sales": resolved_sales,
        "valid_mobiles": valid_mobiles, "sample": sample,
    })
    return res


# -----------------------------------
# 📦 EXPORT FILES — built ONCE per upload (20 xlsx + zip)
# -----------------------------------

try:
    import xlsxwriter  # noqa: F401
    _XLSX_ENGINE = "xlsxwriter"
except ImportError:
    _XLSX_ENGINE = "openpyxl"


def _to_xlsx_bytes(frame):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=_XLSX_ENGINE) as writer:
        frame.to_excel(writer, sheet_name="Sheet1", index=False, startrow=0)
    return buf.getvalue()


@st.cache_resource(show_spinner="📦 Building 20 Excel files & ZIP (runs once per upload)...")
def build_exports(token, _res):
    sheets, dropped_count = build_all_sheets(
        _res["df"], _res["bss_df"], _res["sales_df"], _res["mobile_col_f1"],
        _res["invested_col_f1"], _res["mobile_set_sales"], _res["mobile_set_bss"],
    )

    files, splits, summary_rows = {}, {}, []
    zip_buf = BytesIO()
    total_files = 0
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for _, full_title in SHEET_META:
            data = sheets[full_title]
            tamil = data[data["Language"] == "Tamil"].reset_index(drop=True)
            english = data[data["Language"] != "Tamil"].reset_index(drop=True)
            splits[full_title] = (tamil, english)
            summary_rows.append({
                "Report": full_title,
                "Tamil Rows": len(tamil),
                "English Rows": len(english),
                "Total Rows": len(data),
            })
            for label, subset in (("Tamil", tamil), ("English", english)):
                if len(subset) == 0:
                    continue
                b = _to_xlsx_bytes(subset)
                files[(full_title, label)] = b
                zf.writestr(f"{full_title} - {label}.xlsx", b)
                total_files += 1

    return {
        "dropped_count": dropped_count,
        "files": files,
        "splits": splits,
        "summary": pd.DataFrame(summary_rows),
        "zip_bytes": zip_buf.getvalue(),
        "total_files": total_files,
    }


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
    with st.expander("📄 Customer Report columns", expanded=False):
        st.markdown("- **Permanent Address** *(required)*\n- Temporary Address\n"
                    "- Nearby Showroom\n- Mobile column (any common name) — enables fallback lookup\n"
                    "- Invested column — enables Sheet 2, 7, 8, 9, 10 (Invested splitting)")
    with st.expander("🧾 Bss Receipt / Sales Purchase columns", expanded=False):
        st.markdown("- `Sr` `Date` `Mobileno` `District` `Language`\n"
                    "- District → State → Language auto-filled if Language is blank\n"
                    "- Unmatched sheets pull District/State/Language from their own file row")
    st.divider()
    st.caption("📱 Phone cleaning follows reference logic: strips spaces/hyphens/brackets/plus, "
               "removes letters, +91 → 10-digit (only if valid mobile), leading zeros stripped.")


# -----------------------------------
# Uploaders
# -----------------------------------

st.markdown('<div class="section-title"><span class="step-badge">1</span>Upload Files</div>',
            unsafe_allow_html=True)
col_u1, col_u2, col_u3 = st.columns(3)
with col_u1:
    with st.container(border=True):
        st.markdown("**👥 Customer Report** · required")
        uploaded_file = st.file_uploader("Customer Report *", type=["xlsx", "csv"], key="f1",
                                         label_visibility="collapsed")
with col_u2:
    with st.container(border=True):
        st.markdown("**🧾 Bss Receipt** · optional")
        uploaded_bss = st.file_uploader("Bss Receipt (optional)", type=["xlsx", "csv"], key="f2",
                                        label_visibility="collapsed")
with col_u3:
    with st.container(border=True):
        st.markdown("**🛒 Sales Purchase** · optional")
        uploaded_sales = st.file_uploader("Sales Purchase (optional)", type=["xlsx", "csv"], key="f3",
                                          label_visibility="collapsed")

if not uploaded_file:
    st.info("👋 Upload the Customer Report to get started. Bss Receipt and Sales Purchase are optional.")
    st.stop()


# -----------------------------------
# Run (cached) pipeline
# -----------------------------------

cust_bytes = uploaded_file.getvalue()
bss_bytes = uploaded_bss.getvalue() if uploaded_bss else None
sales_bytes = uploaded_sales.getvalue() if uploaded_sales else None
cust_name = uploaded_file.name
bss_name = uploaded_bss.name if uploaded_bss else ""
sales_name = uploaded_sales.name if uploaded_sales else ""

_h = hashlib.md5()
for _part in (cust_name.encode(), cust_bytes, bss_name.encode(), bss_bytes or b"",
              sales_name.encode(), sales_bytes or b""):
    _h.update(_part)
    _h.update(b"|")
token = _h.hexdigest()

res = run_pipeline(token, cust_bytes, cust_name, bss_bytes, bss_name, sales_bytes, sales_name)

if res["fatal"]:
    st.error(res["fatal"])
    st.stop()

for w in res["warnings"]:
    st.error(w)

df = res["df"]
mobile_col_f1 = res["mobile_col_f1"]
invested_col_f1 = res["invested_col_f1"]
stats_bss, stats_sales = res["stats_bss"], res["stats_sales"]
resolved_bss, resolved_sales = res["resolved_bss"], res["resolved_sales"]
total_resolved = resolved_bss + resolved_sales

st.markdown('<div class="section-title"><span class="step-badge">2</span>Results</div>',
            unsafe_allow_html=True)

if not mobile_col_f1:
    st.warning("⚠️ No mobile column found in the Customer Report — mobile fallback skipped.")
else:
    st.caption(f"📱 Mobile column in Customer Report: `{mobile_col_f1}` — "
               f"{res['valid_mobiles']:,} valid 10-digit mobiles after normalization")

with st.expander("🔎 Mobile Normalization Check (verify +91 → 10 digit)", expanded=False):
    if mobile_col_f1:
        st.dataframe(res["sample"], hide_index=True, width="stretch")
    else:
        st.info("No mobile column detected in Customer Report.")

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
        with st.container(border=True):
            st.markdown("**Language Distribution**")
            st.bar_chart(df["Language"].value_counts())
    with cb:
        with st.container(border=True):
            st.markdown("**Top 10 States**")
            st.bar_chart(df["State"].value_counts().head(10))

MAX_PREVIEW_ROWS = 2000  # keeps the browser fast on very large files (export is never limited)

with tab_results:
    st.subheader("Customer Result")
    fc1, fc2, fc3 = st.columns([2, 1, 1])
    with fc1:
        search_term = st.text_input("🔎 Search address", placeholder="Type part of an address...")
    with fc2:
        state_filter = st.selectbox("Filter by state", ["All"] + sorted(df["State"].unique().tolist()))
    with fc3:
        language_filter = st.selectbox("Filter by language", ["All"] + sorted(df["Language"].unique().tolist()))

    fdf = df
    if search_term:
        fdf = fdf[fdf["Final Address"].str.contains(search_term, case=False, na=False)]
    if state_filter != "All":
        fdf = fdf[fdf["State"] == state_filter]
    if language_filter != "All":
        fdf = fdf[fdf["Language"] == language_filter]

    display_cols = [c for c in ["_Mobile_Key", "Permanent Address", "Temporary Address",
                                "Nearby Showroom", "Final Address", "State", "_District",
                                "Language", "Template", "_MatchedFrom"] if c in fdf.columns]
    st.caption(f"Showing {min(len(fdf), MAX_PREVIEW_ROWS):,} of {len(fdf):,} matching rows "
               f"({total_customers:,} customers total)"
               + (" — preview is capped for speed; exports contain every row." if len(fdf) > MAX_PREVIEW_ROWS else ""))
    st.dataframe(
        fdf[display_cols].head(MAX_PREVIEW_ROWS).rename(columns={"_Mobile_Key": "Mobileno",
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

    exp = build_exports(token, res)

    if exp["dropped_count"] > 0:
        st.info(f"🧹 **Clean-up notice:** {exp['dropped_count']:,} rows with blank/empty phone numbers "
                f"were automatically removed from the final report.")

    if invested_col_f1:
        st.caption(f"Sheets 2, 7, 8, 9, 10 use Customer Report column **`{invested_col_f1}`** — rows where value is `0` or blank are treated as *No Invested*.")
    else:
        st.caption("No invested column found in Customer Report — fallback used: *Invested* = mobiles present in Sales or Bss; *No Invested* = not present in either.")

    st.info("ℹ️ **Sheets 3–6 are rows from the Sales/Bss files themselves.** "
            "District/State/Language on those rows come from that row's own data — "
            "so unmatched rows still display their own District.")

    st.divider()

    # ---- Summary ----
    st.markdown("### 📊 Export Summary (20 Files)")
    st.dataframe(exp["summary"], width="stretch", hide_index=True)

    st.divider()

    # ---- ZIP ----
    st.markdown("### 📦 Download All 20 Files in One ZIP")
    st.caption("The ZIP contains every report split by language: "
               "`<Report> - Tamil.xlsx` and `<Report> - English.xlsx`.")

    st.download_button(
        label=f"📦 Download All {exp['total_files']} Files (.zip)",
        data=exp["zip_bytes"],
        file_name="All_Customer_Reports_Tamil_And_English.zip",
        mime="application/zip",
        key="dl_all_zip",
        width="stretch",
        type="primary",
    )
    st.caption(f"ZIP contains **{exp['total_files']}** Excel files "
               f"(empty language subsets are skipped).")

    st.divider()

    # ---- Individual files ----
    st.markdown("### 📥 Download Individual Files")

    tabs = st.tabs([meta[1] for meta in SHEET_META])
    XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    for tab, (_, full_title) in zip(tabs, SHEET_META):
        tamil_data, english_data = exp["splits"][full_title]
        total_rows = len(tamil_data) + len(english_data)

        with tab:
            st.markdown(f"**{full_title}** — {total_rows:,} total rows "
                        f"({len(tamil_data):,} Tamil • {len(english_data):,} English)")

            col_t, col_e = st.columns(2)

            for col, label, icon, subset in (
                (col_t, "Tamil", "🇮🇳", tamil_data),
                (col_e, "English", "🌍", english_data),
            ):
                with col:
                    with st.container(border=True):
                        st.markdown(f"#### {icon} `{full_title} - {label}.xlsx`")
                        if len(subset) > 0:
                            st.download_button(
                                label=f"📥 Download {label}.xlsx ({len(subset):,} rows)",
                                data=exp["files"][(full_title, label)],
                                file_name=f"{full_title} - {label}.xlsx",
                                mime=XLSX_MIME,
                                key=f"dl_{label.lower()}_{full_title}",
                                width="stretch",
                            )
                            with st.expander(f"Preview {label}", expanded=False):
                                st.dataframe(subset.head(20), hide_index=True, width="stretch")
                        else:
                            st.info(f"No {label} rows in this report.")
                            st.button(
                                f"📥 Download {label}.xlsx (0 rows)",
                                disabled=True,
                                key=f"dl_{label.lower()}_empty_{full_title}",
                                width="stretch",
                            )