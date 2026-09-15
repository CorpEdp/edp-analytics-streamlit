"""
Auto-converted from: App-CustomerRegisterLocationFind.py
Review this file before using — the converter does a best-effort wrap;
double-check indentation around any unusual control flow (loops, if/else
blocks that span large sections, etc.).
"""

import streamlit as st
import pandas as pd
from io import BytesIO
import re
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
    text = str(text)
    match = re.search(r'\b\d{6}\b', text)
    if match:
        return int(match.group())
    return None
def get_state_from_pincode(pincode):
    """Return the state for a pincode, preferring the narrowest (most
    specific) matching range when multiple ranges overlap."""
    if not pincode:
        return None
    best_state = None
    best_width = None
    for start, end, state in PINCODE_RANGES:
        if start <= pincode <= end:
            width = end - start
            if best_width is None or width < best_width:
                best_width = width
                best_state = state
    return best_state
def get_state_from_variation(text):
    clean = clean_text(text)
    tn_variations = [
        "tamilnadu", "tamil nadu", "tamilnad", "tamailnadu",
        "tanilnadu", "thamilnadu", "thamil nadu",
        "thamiznadu", "thamizh nadu", "thamizhnadu",
        "tamizhnadu", "tamizh nadu",
        "taminadu", "taminad", "tamnad", "tamnadu",
        "thamil natu", "tamil natu",
        "tami nadu", "tamil nadu",
        "tn", "tamil"
    ]
    for variation in tn_variations:
        if variation in clean:
            return "Tamil Nadu"

    for variation, state in state_variations.items():
        if variation in clean:
            return state
    return None
def fuzzy_state_match(text):
    """Handle common misspellings using partial matching"""
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
    perm_addr = row.get("Permanent Address", "")
    if pd.notna(perm_addr) and str(perm_addr).strip():
        return str(perm_addr).strip()

    temp_addr = row.get("Temporary Address", "")
    if pd.notna(temp_addr) and str(temp_addr).strip():
        return str(temp_addr).strip()

    nearby = row.get("Nearby Showroom", "")
    if pd.notna(nearby) and str(nearby).strip():
        return str(nearby).strip()

    return ""
def find_location(address):
    if pd.isna(address) or str(address).strip() == "":
        return "Unknown", "English"

    original = str(address)
    clean_address = clean_text(original)

    if not clean_address:
        return "Unknown", "English"

    explicit_state_keywords = {
        "kerala": "Kerala",
        "karnataka": "Karnataka",
        "tamil nadu": "Tamil Nadu",
        "tamilnadu": "Tamil Nadu",
        "andhra pradesh": "Andhra Pradesh",
        "andhrapradesh": "Andhra Pradesh",
        "telangana": "Telangana",
        "maharashtra": "Maharashtra",
        "gujarat": "Gujarat",
        "west bengal": "West Bengal",
        "westbengal": "West Bengal",
        "odisha": "Odisha",
        "orissa": "Odisha",
        "uttar pradesh": "Uttar Pradesh",
        "uttarpradesh": "Uttar Pradesh",
        "bihar": "Bihar",
        "rajasthan": "Rajasthan",
        "madhya pradesh": "Madhya Pradesh",
        "madhyapradesh": "Madhya Pradesh",
        "delhi": "Delhi",
        "punjab": "Punjab",
        "haryana": "Haryana",
        "jharkhand": "Jharkhand",
        "chhattisgarh": "Chhattisgarh",
        "chattisgarh": "Chhattisgarh",
        "assam": "Assam",
        "goa": "Goa",
        "puducherry": "Puducherry",
        "pondicherry": "Puducherry",
        "jammu & kashmir": "Jammu & Kashmir",
        "jammu and kashmir": "Jammu & Kashmir",
        "himachal pradesh": "Himachal Pradesh",
        "himachalpradesh": "Himachal Pradesh",
        "uttarakhand": "Uttarakhand",
        "uttaranchal": "Uttarakhand",
    }

    for keyword, state in explicit_state_keywords.items():
        if keyword in clean_address:
            return state, state_language.get(state, "English")

    tn_locations = [
        "ariyalur", "chengalpattu", "chennai",
        "coimbatore", "coimpatore", "coimbtore", "coimbatur",
        "coimbathore", "coyambatore", "coyambathore", "kovai",
        "cuddalore", "dharmapuri", "dindigul", "erode",
        "kallakurichi", "kanchipuram", "kanyakumari", "kanniyakumari",
        "kanaya kumari", "kaniya kumari", "karur", "krishnagiri",
        "madurai", "mayiladuthurai", "nagapattinam",
        "namakkal", "nammakkal", "namakal",
        "nilgiris", "perambalur", "pudukkottai", "ramanathapuram",
        "ranipet", "salem", "sivaganga", "sivagangai", "sivgangai",
        "sivagngai", "sivagankai", "tenkasi", "thanjavur", "theni",
        "thoothukudi", "thuthukudi", "tutukkudi",
        "tutukudi", "tuthukudi", "tuttukudi",
        "tuticorin", "tiruchirappalli",
        "thiruchirapalli", "trichy", "tirunelveli", "thirunelveli",
        "tirupathur", "tiruppur", "tirupur",
        "tiruvallur", "tiruvannamalai", "tiruvarur", "vellore",
        "viluppuram", "virudhunagar", "viruthunagar", "viridhunagar",
        "virudunagar", "vridhunagar", "varudhu nagar", "varudhunagar",
        "karaikudi", "karaikudy", "karaikkudi",
        "ooty", "hosur", "sivakasi", "rajapalayam",
        "nagercoil", "marthandam", "kollamcode", "kollamcodu",
        "kollamkode", "kaliyakkavilai", "kulasekharam", "kuzhithurai",
        "thuckalay", "aruppukottai", "sivagiri", "palayamkottai",
        "tirumangalam", "thirumangalam", "usilampatti", "thenkasi",
        "sankarankovil", "kadayanallur", "ambasamudram", "cheranmahadevi",
        "nanguneri", "valliyur", "thisayanvilai", "kovilpatti", "sattur",
        "srivaikundam", "kilakarai", "paramakudi", "vaniyambadi",
        "gudiyattam", "arni", "polur", "sriperumbudur", "tambaram",
        "avadi", "pallavaram", "chromepet", "poonamallee", "ambattur",
        "madhuravoyal", "porur", "perungudi", "mylapore", "adyar",
        "anna nagar", "t nagar", "egmore", "nungambakkam", "kodambakkam",
        "perambur", "pulianthope", "george town", "parrys", "broadway",
        "washermanpet", "tondiarpet",
        "kirthoor", "kollemcode", "thiruvattar", "marthadam", "kunnathoor",
        "kunnathur", "kanniya", "puthukkadai", "veeyannoor",
        "kallidaikkurichchi", "kallidaikurichi",
        "melapalayam", "pattenkaliuru", "alangulam", "manamadurai",
        "attyapatti", "rasipuram", "kundrakudi", "kallal", "liyankudi",
        "sivangai", "ramachandrapuram", "thagapattanam", "thagapattinam",
        "seyganallur", "seiganallur", "uthappanaickanur", "uthappanaickenur",
        "thirupuvanam", "tirupuvanam", "kalpakk", "kalpak", "watrap",
        "thirachangodu", "thirachangode", "pulavanvilai", "pulavanvila",
        "kalankuzhi", "murambu"
    ]

    for location in tn_locations:
        if location in clean_address:
            return "Tamil Nadu", "Tamil"

    py_locations = ["puducherry", "pondicherry", "pondocherry", "karaikal",
                    "thattanchavady", "villianur", "iyyankuttipalayam",
                    "bahour", "mannadipet", "nettapakkam", "ariyankuppam"]
    for location in py_locations:
        if location in clean_address:
            return "Puducherry", "Tamil"

    tn_variations = [
        "tamilnadu", "tamil nadu", "tamilnad", "tamailnadu",
        "tanilnadu", "thamilnadu", "thamil nadu",
        "thamiznadu", "thamizh nadu", "thamizhnadu",
        "tamizhnadu", "tamizh nadu",
        "taminadu", "taminad", "tamnad", "tamnadu",
        "thamil natu", "tamil natu",
        "tami nadu", "tn", "tamil"
    ]
    for variation in tn_variations:
        if variation in clean_address:
            return "Tamil Nadu", "Tamil"

    variation_state = get_state_from_variation(clean_address)
    if variation_state:
        return variation_state, state_language.get(variation_state, "English")

    fuzzy_state = fuzzy_state_match(clean_address)
    if fuzzy_state:
        return fuzzy_state, state_language.get(fuzzy_state, "English")

    for city, state in city_state.items():
        pattern = r'\b' + re.escape(city) + r'\b'
        if re.search(pattern, clean_address):
            return state, state_language.get(state, "English")

    for city, state in city_state.items():
        if city in clean_address:
            return state, state_language.get(state, "English")

    pincode = extract_pincode(original)
    if pincode:
        state = get_state_from_pincode(pincode)
        if state:
            return state, state_language.get(state, "English")

    words = clean_address.split()
    for word in words:
        for state in state_language.keys():
            if state.lower() in word or word in state.lower():
                return state, state_language[state]

        for city, state in city_state.items():
            if city in word or word in city:
                return state, state_language.get(state, "English")

    state_abbrs = {
        "tn": "Tamil Nadu", "ka": "Karnataka", "kl": "Kerala",
        "ap": "Andhra Pradesh", "ts": "Telangana", "mh": "Maharashtra",
        "gj": "Gujarat", "wb": "West Bengal", "up": "Uttar Pradesh",
        "br": "Bihar", "rj": "Rajasthan", "mp": "Madhya Pradesh",
        "pb": "Punjab", "hr": "Haryana", "dl": "Delhi",
        "od": "Odisha", "ga": "Goa", "py": "Puducherry",
        "jk": "Jammu & Kashmir", "hp": "Himachal Pradesh", "uk": "Uttarakhand"
    }
    for abbr, state in state_abbrs.items():
        if abbr in clean_address:
            return state, state_language.get(state, "English")

    return "Other", "English"

LABEL = "Customer Register Location Find"


def run():
    st.title("📍 Customer State & Language Finder")
    st.caption("Upload a customer report and automatically detect each customer's state, "
               "regional language, and messaging template from their address.")
    state_language = {
        "Tamil Nadu": "Tamil",
        "Karnataka": "Kannada",
        "Kerala": "Malayalam",
        "Andhra Pradesh": "Telugu",
        "Telangana": "Telugu",
        "Maharashtra": "Marathi",
        "Gujarat": "Gujarati",
        "West Bengal": "Bengali",
        "Odisha": "Odia",
        "Punjab": "Punjabi",
        "Rajasthan": "Hindi",
        "Uttar Pradesh": "Hindi",
        "Bihar": "Hindi",
        "Madhya Pradesh": "Hindi",
        "Delhi": "Hindi",
        "Haryana": "Hindi",
        "Jharkhand": "Hindi",
        "Chhattisgarh": "Hindi",
        "Assam": "Assamese",
        "Goa": "Konkani",
        "Manipur": "Manipuri",
        "Mizoram": "Mizo",
        "Sikkim": "Nepali",
        "Tripura": "Bengali",
        "Puducherry": "Tamil",
        "Pondicherry": "Tamil",
        "Jammu & Kashmir": "Urdu/Hindi",
        "Jammu and Kashmir": "Urdu/Hindi",
        "Ladakh": "Ladakhi",
        "Himachal Pradesh": "Hindi",
        "Uttarakhand": "Hindi",
    }
    state_variations = {
        # Tamil Nadu - COMPLETE with ALL variations
        "tamilnadu": "Tamil Nadu",
        "tamil nadu": "Tamil Nadu",
        "tamilnad": "Tamil Nadu",
        "tamailnadu": "Tamil Nadu",
        "tanilnadu": "Tamil Nadu",
        "thamilnadu": "Tamil Nadu",
        "thamil nadu": "Tamil Nadu",
        "thamiznadu": "Tamil Nadu",
        "thamizh nadu": "Tamil Nadu",
        "thamizhnadu": "Tamil Nadu",
        "tamizhnadu": "Tamil Nadu",
        "tamizh nadu": "Tamil Nadu",
        "taminadu": "Tamil Nadu",
        "taminad": "Tamil Nadu",
        "tamnad": "Tamil Nadu",
        "tamnadu": "Tamil Nadu",
        "tami nadu": "Tamil Nadu",
        "thamil natu": "Tamil Nadu",
        "tamil natu": "Tamil Nadu",
        "tn": "Tamil Nadu",
        "tamil": "Tamil Nadu",
        "madras": "Tamil Nadu",
        "chennai": "Tamil Nadu",

        # Karnataka
        "karnataka": "Karnataka",
        "karnatka": "Karnataka",
        "karnatak": "Karnataka",
        "ka": "Karnataka",
        "bangalore": "Karnataka",
        "bengaluru": "Karnataka",

        # Kerala
        "kerala": "Kerala",
        "kerla": "Kerala",
        "kerela": "Kerala",
        "keral": "Kerala",
        "kerlaa": "Kerala",
        "keralam": "Kerala",
        "kl": "Kerala",
        "kochi": "Kerala",
        "trivandrum": "Kerala",

        # Andhra Pradesh
        "andhra pradesh": "Andhra Pradesh",
        "andhra": "Andhra Pradesh",
        "andhrapradesh": "Andhra Pradesh",
        "ap": "Andhra Pradesh",

        # Telangana
        "telangana": "Telangana",
        "telangna": "Telangana",
        "ts": "Telangana",
        "hyderabad": "Telangana",

        # Maharashtra
        "maharashtra": "Maharashtra",
        "maharastra": "Maharashtra",
        "maharashstra": "Maharashtra",
        "mh": "Maharashtra",
        "mumbai": "Maharashtra",
        "pune": "Maharashtra",

        # Gujarat
        "gujarat": "Gujarat",
        "gujrat": "Gujarat",
        "gj": "Gujarat",
        "ahmedabad": "Gujarat",

        # West Bengal
        "west bengal": "West Bengal",
        "westbengal": "West Bengal",
        "wb": "West Bengal",
        "kolkata": "West Bengal",

        # Odisha
        "odisha": "Odisha",
        "orissa": "Odisha",
        "orrisa": "Odisha",
        "od": "Odisha",

        # Uttar Pradesh
        "uttar pradesh": "Uttar Pradesh",
        "uttarpradesh": "Uttar Pradesh",
        "up": "Uttar Pradesh",
        "lucknow": "Uttar Pradesh",

        # Bihar
        "bihar": "Bihar",
        "br": "Bihar",
        "patna": "Bihar",

        # Rajasthan
        "rajasthan": "Rajasthan",
        "rj": "Rajasthan",
        "jaipur": "Rajasthan",

        # Madhya Pradesh
        "madhya pradesh": "Madhya Pradesh",
        "madhyapradesh": "Madhya Pradesh",
        "mp": "Madhya Pradesh",
        "bhopal": "Madhya Pradesh",

        # Delhi
        "delhi": "Delhi",
        "dl": "Delhi",

        # Punjab
        "punjab": "Punjab",
        "pb": "Punjab",
        "chandigarh": "Punjab",

        # Haryana
        "haryana": "Haryana",
        "haryan": "Haryana",
        "hr": "Haryana",

        # Jharkhand
        "jharkhand": "Jharkhand",
        "jh": "Jharkhand",

        # Chhattisgarh
        "chhattisgarh": "Chhattisgarh",
        "chattisgarh": "Chhattisgarh",
        "cg": "Chhattisgarh",

        # Assam
        "assam": "Assam",
        "as": "Assam",

        # Goa
        "goa": "Goa",
        "ga": "Goa",

        # Puducherry
        "puducherry": "Puducherry",
        "pondicherry": "Puducherry",
        "py": "Puducherry",

        # Jammu & Kashmir
        "jammu & kashmir": "Jammu & Kashmir",
        "jammu and kashmir": "Jammu & Kashmir",
        "jammu&kashmir": "Jammu & Kashmir",
        "jammukashmir": "Jammu & Kashmir",
        "j&k": "Jammu & Kashmir",
        "jk": "Jammu & Kashmir",
        "jammu": "Jammu & Kashmir",
        "kashmir": "Jammu & Kashmir",

        # Himachal Pradesh
        "himachal pradesh": "Himachal Pradesh",
        "himachalpradesh": "Himachal Pradesh",
        "hp": "Himachal Pradesh",
        "himachal": "Himachal Pradesh",

        # Uttarakhand
        "uttarakhand": "Uttarakhand",
        "uttaranchal": "Uttarakhand",
        "uk": "Uttarakhand",
    }
    city_state = {
        # ===== TAMIL NADU - ALL DISTRICTS AND MAJOR CITIES =====
        "ariyalur": "Tamil Nadu",
        "chengalpattu": "Tamil Nadu",
        "chennai": "Tamil Nadu",
        "coimbatore": "Tamil Nadu",
        "coimpatore": "Tamil Nadu",
        "coimbtore": "Tamil Nadu",
        "coimbatur": "Tamil Nadu",
        "coimbathore": "Tamil Nadu",
        "coyambatore": "Tamil Nadu",
        "coyambathore": "Tamil Nadu",
        "kovai": "Tamil Nadu",
        "cbe": "Tamil Nadu",
        "cuddalore": "Tamil Nadu",
        "dharmapuri": "Tamil Nadu",
        "dindigul": "Tamil Nadu",
        "erode": "Tamil Nadu",
        "kallakurichi": "Tamil Nadu",
        "kanchipuram": "Tamil Nadu",
        "kanyakumari": "Tamil Nadu",
        "kanniyakumari": "Tamil Nadu",
        "kanaya kumari": "Tamil Nadu",
        "kaniya kumari": "Tamil Nadu",
        "kaniyakumari": "Tamil Nadu",
        "karur": "Tamil Nadu",
        "krishnagiri": "Tamil Nadu",
        "madurai": "Tamil Nadu",
        "mayiladuthurai": "Tamil Nadu",
        "nagapattinam": "Tamil Nadu",
        "namakkal": "Tamil Nadu",
        "nammakkal": "Tamil Nadu",
        "namakal": "Tamil Nadu",
        "nilgiris": "Tamil Nadu",
        "perambalur": "Tamil Nadu",
        "pudukkottai": "Tamil Nadu",
        "ramanathapuram": "Tamil Nadu",
        "ranipet": "Tamil Nadu",
        "salem": "Tamil Nadu",
        "sivaganga": "Tamil Nadu",
        "sivagangai": "Tamil Nadu",
        "sivgangai": "Tamil Nadu",
        "sivagngai": "Tamil Nadu",
        "sivagankai": "Tamil Nadu",
        "tenkasi": "Tamil Nadu",
        "thanjavur": "Tamil Nadu",
        "theni": "Tamil Nadu",
        "thoothukudi": "Tamil Nadu",
        "thuthukudi": "Tamil Nadu",
        "tutukkudi": "Tamil Nadu",
        "tutukudi": "Tamil Nadu",
        "tuthukudi": "Tamil Nadu",
        "tuttukudi": "Tamil Nadu",
        "tuticorin": "Tamil Nadu",
        "tiruchirappalli": "Tamil Nadu",
        "thiruchirapalli": "Tamil Nadu",
        "trichy": "Tamil Nadu",
        "tirunelveli": "Tamil Nadu",
        "thirunelveli": "Tamil Nadu",
        "tirupathur": "Tamil Nadu",
        "tiruppur": "Tamil Nadu",
        "tirupur": "Tamil Nadu",
        "tiruvallur": "Tamil Nadu",
        "tiruvannamalai": "Tamil Nadu",
        "tiruvarur": "Tamil Nadu",
        "vellore": "Tamil Nadu",
        "viluppuram": "Tamil Nadu",
        "virudhunagar": "Tamil Nadu",
        "viruthunagar": "Tamil Nadu",
        "viridhunagar": "Tamil Nadu",
        "virudunagar": "Tamil Nadu",
        "vridhunagar": "Tamil Nadu",
        "varudhu nagar": "Tamil Nadu",
        "varudhunagar": "Tamil Nadu",

        # Cities and Towns
        "karaikudi": "Tamil Nadu",
        "karaikudy": "Tamil Nadu",
        "karaikkudi": "Tamil Nadu",
        "ooty": "Tamil Nadu",
        "hosur": "Tamil Nadu",
        "sivakasi": "Tamil Nadu",
        "rajapalayam": "Tamil Nadu",
        "nagercoil": "Tamil Nadu",
        "marthandam": "Tamil Nadu",
        "kollamcode": "Tamil Nadu",
        "kollamcodu": "Tamil Nadu",
        "kollamkode": "Tamil Nadu",
        "kaliyakkavilai": "Tamil Nadu",
        "kulasekharam": "Tamil Nadu",
        "kuzhithurai": "Tamil Nadu",
        "thuckalay": "Tamil Nadu",
        "aruppukottai": "Tamil Nadu",
        "sivagiri": "Tamil Nadu",
        "palayamkottai": "Tamil Nadu",
        "tirumangalam": "Tamil Nadu",
        "thirumangalam": "Tamil Nadu",
        "usilampatti": "Tamil Nadu",
        "thenkasi": "Tamil Nadu",
        "sankarankovil": "Tamil Nadu",
        "kadayanallur": "Tamil Nadu",
        "ambasamudram": "Tamil Nadu",
        "cheranmahadevi": "Tamil Nadu",
        "nanguneri": "Tamil Nadu",
        "valliyur": "Tamil Nadu",
        "thisayanvilai": "Tamil Nadu",
        "kovilpatti": "Tamil Nadu",
        "sattur": "Tamil Nadu",
        "srivaikundam": "Tamil Nadu",
        "kilakarai": "Tamil Nadu",
        "paramakudi": "Tamil Nadu",
        "vaniyambadi": "Tamil Nadu",
        "gudiyattam": "Tamil Nadu",
        "arni": "Tamil Nadu",
        "polur": "Tamil Nadu",
        "sriperumbudur": "Tamil Nadu",
        "tambaram": "Tamil Nadu",
        "avadi": "Tamil Nadu",
        "pallavaram": "Tamil Nadu",
        "chromepet": "Tamil Nadu",
        "poonamallee": "Tamil Nadu",
        "ambattur": "Tamil Nadu",
        "madhuravoyal": "Tamil Nadu",
        "porur": "Tamil Nadu",
        "perungudi": "Tamil Nadu",
        "mylapore": "Tamil Nadu",
        "adyar": "Tamil Nadu",
        "anna nagar": "Tamil Nadu",
        "t nagar": "Tamil Nadu",
        "egmore": "Tamil Nadu",
        "nungambakkam": "Tamil Nadu",
        "kodambakkam": "Tamil Nadu",
        "perambur": "Tamil Nadu",
        "pulianthope": "Tamil Nadu",
        "george town": "Tamil Nadu",
        "parrys": "Tamil Nadu",
        "broadway": "Tamil Nadu",
        "washermanpet": "Tamil Nadu",
        "tondiarpet": "Tamil Nadu",
        "thagapattanam": "Tamil Nadu",
        "thagapattinam": "Tamil Nadu",
        "seyganallur": "Tamil Nadu",
        "seiganallur": "Tamil Nadu",
        "uthappanaickanur": "Tamil Nadu",
        "uthappanaickenur": "Tamil Nadu",
        "thirupuvanam": "Tamil Nadu",
        "tirupuvanam": "Tamil Nadu",
        "kalpakk": "Tamil Nadu",
        "kalpak": "Tamil Nadu",
        "watrap": "Tamil Nadu",
        "thirachangodu": "Tamil Nadu",
        "thirachangode": "Tamil Nadu",
        "pulavanvilai": "Tamil Nadu",
        "pulavanvila": "Tamil Nadu",
        "kalankuzhi": "Tamil Nadu",
        "murambu": "Tamil Nadu",

        # KANNIYAKUMARI DISTRICT
        "kirthoor": "Tamil Nadu",
        "kollemcode": "Tamil Nadu",
        "thiruvattar": "Tamil Nadu",
        "marthadam": "Tamil Nadu",
        "kunnathoor": "Tamil Nadu",
        "kunnathur": "Tamil Nadu",
        "kanniya": "Tamil Nadu",
        "puthukkadai": "Tamil Nadu",
        "veeyannoor": "Tamil Nadu",
        "ramachandrapuram": "Tamil Nadu",
        "kallidaikkurichchi": "Tamil Nadu",
        "kallidaikurichi": "Tamil Nadu",

        "melapalayam": "Tamil Nadu",
        "pattenkaliuru": "Tamil Nadu",
        "alangulam": "Tamil Nadu",
        "manamadurai": "Tamil Nadu",
        "attyapatti": "Tamil Nadu",
        "rasipuram": "Tamil Nadu",
        "kundrakudi": "Tamil Nadu",
        "kallal": "Tamil Nadu",
        "liyankudi": "Tamil Nadu",
        "sivangai": "Tamil Nadu",

        # ===== KERALA =====
        "kochi": "Kerala",
        "cochin": "Kerala",
        "trivandrum": "Kerala",
        "thiruvananthapuram": "Kerala",
        "kozhikode": "Kerala",
        "calicut": "Kerala",
        "thrissur": "Kerala",
        "trichur": "Kerala",
        "kollam": "Kerala",
        "alappuzha": "Kerala",
        "alleppey": "Kerala",
        "palakkad": "Kerala",
        "palghat": "Kerala",
        "pattambi": "Kerala",
        "pattamby": "Kerala",
        "pattambii": "Kerala",
        "ottapalam": "Kerala",
        "shoranur": "Kerala",
        "kannur": "Kerala",
        "cannanore": "Kerala",
        "kasargod": "Kerala",
        "pathanamthitta": "Kerala",
        "idukki": "Kerala",
        "wayanad": "Kerala",
        "kottayam": "Kerala",
        "ernakulam": "Kerala",
        "malappuram": "Kerala",
        "neyyattinkara": "Kerala",
        "parassala": "Kerala",
        "perinthalmanna": "Kerala",
        "manjeri": "Kerala",
        "ponnani": "Kerala",
        "chavakkad": "Kerala",
        "guruvayur": "Kerala",
        "irijalakkuda": "Kerala",
        "chengannur": "Kerala",
        "mavelikkara": "Kerala",
        "kayamkulam": "Kerala",
        "valakam": "Kerala",
        "valakom": "Kerala",
        "valacom": "Kerala",

        # ===== PUDUCHERRY =====
        "puducherry": "Puducherry",
        "pondicherry": "Puducherry",
        "pondocherry": "Puducherry",
        "karaikal": "Puducherry",
        "thattanchavady": "Puducherry",
        "villianur": "Puducherry",
        "iyyankuttipalayam": "Puducherry",
        "bahour": "Puducherry",
        "mannadipet": "Puducherry",
        "nettapakkam": "Puducherry",
        "ariyankuppam": "Puducherry",
        "embalam": "Puducherry",

        # ===== KARNATAKA =====
        "bangalore": "Karnataka",
        "bengaluru": "Karnataka",
        "mysore": "Karnataka",
        "mysuru": "Karnataka",
        "mangalore": "Karnataka",
        "mangaluru": "Karnataka",
        "hubli": "Karnataka",
        "dharwad": "Karnataka",
        "belgaum": "Karnataka",
        "belagavi": "Karnataka",
        "gulbarga": "Karnataka",
        "kalaburagi": "Karnataka",
        "bijapur": "Karnataka",
        "vijayapura": "Karnataka",
        "bellary": "Karnataka",
        "ballari": "Karnataka",
        "tumkur": "Karnataka",
        "tumakuru": "Karnataka",
        "shimoga": "Karnataka",
        "shivamogga": "Karnataka",
        "davanagere": "Karnataka",
        "raichur": "Karnataka",
        "bagalkot": "Karnataka",
        "hassan": "Karnataka",
        "kolar": "Karnataka",
        "mandya": "Karnataka",
        "chikkamagaluru": "Karnataka",
        "udupi": "Karnataka",
        "chikmagalur": "Karnataka",
        "chitradurga": "Karnataka",
        "hospet": "Karnataka",
        "bhadravati": "Karnataka",
        "gangavathi": "Karnataka",
        "gadag": "Karnataka",
        "bidar": "Karnataka",

        # ===== ANDHRA PRADESH =====
        "vijayawada": "Andhra Pradesh",
        "visakhapatnam": "Andhra Pradesh",
        "vizag": "Andhra Pradesh",
        "guntur": "Andhra Pradesh",
        "nellore": "Andhra Pradesh",
        "kurnool": "Andhra Pradesh",
        "tirupati": "Andhra Pradesh",
        "rajahmundry": "Andhra Pradesh",
        "kakinada": "Andhra Pradesh",
        "chittoor": "Andhra Pradesh",
        "anantapur": "Andhra Pradesh",
        "kadapa": "Andhra Pradesh",
        "hindupur": "Andhra Pradesh",
        "machilipatnam": "Andhra Pradesh",
        "tenali": "Andhra Pradesh",
        "nandyal": "Andhra Pradesh",

        # ===== TELANGANA =====
        "hyderabad": "Telangana",
        "secunderabad": "Telangana",
        "warangal": "Telangana",
        "nizamabad": "Telangana",
        "karimnagar": "Telangana",
        "khammam": "Telangana",
        "ramagundam": "Telangana",
        "mahabubnagar": "Telangana",
        "nalgonda": "Telangana",
        "siddipet": "Telangana",
        "miryalaguda": "Telangana",
        "jagtial": "Telangana",

        # ===== MAHARASHTRA =====
        "mumbai": "Maharashtra",
        "pune": "Maharashtra",
        "nagpur": "Maharashtra",
        "thane": "Maharashtra",
        "nashik": "Maharashtra",
        "aurangabad": "Maharashtra",
        "solapur": "Maharashtra",
        "kalyan": "Maharashtra",
        "vasai": "Maharashtra",
        "navi mumbai": "Maharashtra",

        # ===== GUJARAT =====
        "ahmedabad": "Gujarat",
        "surat": "Gujarat",
        "vadodara": "Gujarat",
        "rajkot": "Gujarat",
        "bhavnagar": "Gujarat",
        "jamnagar": "Gujarat",
        "junagadh": "Gujarat",
        "gandhinagar": "Gujarat",
        "anand": "Gujarat",
        "nadiad": "Gujarat",
        "morbi": "Gujarat",
        "bhuj": "Gujarat",
        "gandhidham": "Gujarat",

        # ===== WEST BENGAL =====
        "kolkata": "West Bengal",
        "howrah": "West Bengal",
        "durgapur": "West Bengal",
        "siliguri": "West Bengal",
        "asansol": "West Bengal",

        # ===== DELHI NCR =====
        "delhi": "Delhi",
        "new delhi": "Delhi",
        "gurugram": "Haryana",
        "gurgaon": "Haryana",
        "noida": "Uttar Pradesh",
        "ghaziabad": "Uttar Pradesh",
        "faridabad": "Haryana",
        "greater noida": "Uttar Pradesh",
        "sahibabad": "Uttar Pradesh",
        "indirapuram": "Uttar Pradesh",

        # ===== UTTAR PRADESH =====
        "lucknow": "Uttar Pradesh",
        "kanpur": "Uttar Pradesh",
        "varanasi": "Uttar Pradesh",
        "agra": "Uttar Pradesh",
        "allahabad": "Uttar Pradesh",
        "prayagraj": "Uttar Pradesh",
        "meerut": "Uttar Pradesh",
        "aligarh": "Uttar Pradesh",
        "bareilly": "Uttar Pradesh",
        "moradabad": "Uttar Pradesh",
        "saharanpur": "Uttar Pradesh",

        # ===== BIHAR =====
        "patna": "Bihar",
        "gaya": "Bihar",
        "bhagalpur": "Bihar",
        "muzaffarpur": "Bihar",
        "darbhanga": "Bihar",
        "siwan": "Bihar",
        "sasaram": "Bihar",
        "ara": "Bihar",

        # ===== RAJASTHAN =====
        "jaipur": "Rajasthan",
        "jodhpur": "Rajasthan",
        "udaipur": "Rajasthan",
        "kota": "Rajasthan",
        "bikaner": "Rajasthan",
        "ajmer": "Rajasthan",
        "chittorgarh": "Rajasthan",

        # ===== MADHYA PRADESH =====
        "bhopal": "Madhya Pradesh",
        "indore": "Madhya Pradesh",
        "gwalior": "Madhya Pradesh",
        "jabalpur": "Madhya Pradesh",
        "ujjain": "Madhya Pradesh",
        "sagar": "Madhya Pradesh",
        "dewas": "Madhya Pradesh",
        "ratlam": "Madhya Pradesh",

        # ===== PUNJAB =====
        "chandigarh": "Punjab",
        "amritsar": "Punjab",
        "ludhiana": "Punjab",
        "jalandhar": "Punjab",
        "patiala": "Punjab",
        "mohali": "Punjab",

        # ===== ASSAM =====
        "guwahati": "Assam",
        "dibrugarh": "Assam",
        "silchar": "Assam",
        "jorhat": "Assam",

        # ===== GOA =====
        "panaji": "Goa",
        "margao": "Goa",
        "vasco": "Goa",
        "mapusa": "Goa",
        "ponda": "Goa",

        # ===== ODISHA =====
        "bhubaneswar": "Odisha",
        "cuttack": "Odisha",
        "rourkela": "Odisha",
        "puri": "Odisha",
        "sambalpur": "Odisha",
        "berhampur": "Odisha",
        "balangir": "Odisha",

        # ===== JAMMU & KASHMIR =====
        "jammu": "Jammu & Kashmir",
        "srinagar": "Jammu & Kashmir",
        "kashmir": "Jammu & Kashmir",
        "kargil": "Jammu & Kashmir",
        "leh": "Jammu & Kashmir",
        "ladakh": "Jammu & Kashmir",
        "anantnag": "Jammu & Kashmir",
        "baramulla": "Jammu & Kashmir",
        "pulwama": "Jammu & Kashmir",
        "kupwara": "Jammu & Kashmir",
        "badgam": "Jammu & Kashmir",
        "shopian": "Jammu & Kashmir",
        "ganderbal": "Jammu & Kashmir",
        "bandipora": "Jammu & Kashmir",
        "samba": "Jammu & Kashmir",
        "udhampur": "Jammu & Kashmir",
        "reasi": "Jammu & Kashmir",
        "ramban": "Jammu & Kashmir",
        "doda": "Jammu & Kashmir",
        "kishtwar": "Jammu & Kashmir",
        "poonch": "Jammu & Kashmir",
        "rajouri": "Jammu & Kashmir",
        "naushera": "Jammu & Kashmir",
        "akhnoor": "Jammu & Kashmir",
        "kathua": "Jammu & Kashmir",

        # ===== HIMACHAL PRADESH =====
        "shimla": "Himachal Pradesh",
        "manali": "Himachal Pradesh",
        "dharamshala": "Himachal Pradesh",
        "kullu": "Himachal Pradesh",
        "mandi": "Himachal Pradesh",
        "solan": "Himachal Pradesh",
        "kangra": "Himachal Pradesh",
        "chamba": "Himachal Pradesh",
        "hamirpur": "Himachal Pradesh",
        "una": "Himachal Pradesh",
        "bilaspur": "Himachal Pradesh",

        # ===== UTTARAKHAND =====
        "dehradun": "Uttarakhand",
        "haridwar": "Uttarakhand",
        "rishikesh": "Uttarakhand",
        "nainital": "Uttarakhand",
        "mussoorie": "Uttarakhand",
        "almora": "Uttarakhand",
        "pithoragarh": "Uttarakhand",
        "rudrapur": "Uttarakhand",
        "haldwani": "Uttarakhand",
        "roorkee": "Uttarakhand",
    }
    PINCODE_RANGES = [
        # Tamil Nadu (600000-643999), with Puducherry enclaves carved out below
        (600000, 643999, "Tamil Nadu"),

        # Puducherry enclaves (more specific/narrower -> take precedence)
        (605001, 605013, "Puducherry"),   # Puducherry town
        (607401, 607409, "Puducherry"),
        (609601, 609609, "Puducherry"),   # Karaikal

        # Kerala
        (670000, 696999, "Kerala"),

        # Karnataka
        (560000, 591999, "Karnataka"),

        # Andhra Pradesh
        (515000, 534999, "Andhra Pradesh"),

        # Telangana
        (500000, 509999, "Telangana"),

        # Maharashtra
        (400000, 449999, "Maharashtra"),

        # Gujarat
        (360000, 395999, "Gujarat"),

        # West Bengal
        (700000, 700999, "West Bengal"),
        (710000, 744999, "West Bengal"),

        # Delhi
        (110000, 110999, "Delhi"),

        # Haryana
        (120000, 139999, "Haryana"),

        # Uttar Pradesh (200000-247999 and 264000-289999; 248000-263999 belongs
        # to Uttarakhand, which was carved out of UP and is listed separately)
        (200000, 247999, "Uttar Pradesh"),
        (264000, 289999, "Uttar Pradesh"),

        # Uttarakhand (carved out of UP in 2000)
        (244000, 263999, "Uttarakhand"),

        # Rajasthan
        (300000, 349999, "Rajasthan"),

        # Madhya Pradesh
        (450000, 499999, "Madhya Pradesh"),

        # Bihar
        (800000, 859999, "Bihar"),

        # Odisha
        (751000, 770999, "Odisha"),

        # Jammu & Kashmir
        (180000, 194999, "Jammu & Kashmir"),

        # Himachal Pradesh
        (171000, 177999, "Himachal Pradesh"),
    ]
    with st.sidebar:
        st.header("ℹ️ How it works")
        st.markdown(
            "1. Upload a customer report (`.xlsx` or `.csv`)\n"
            "2. The app picks the best available address per customer:\n"
            "   `Permanent Address` → `Temporary Address` → `Nearby Showroom`\n"
            "3. State, language & template are detected from that address\n"
            "4. Review, filter, and download the enriched report"
        )
        st.divider()
        st.subheader("Required columns")
        st.markdown("- **Permanent Address** *(required)*")
        st.markdown("- Temporary Address *(optional)*")
        st.markdown("- Nearby Showroom *(optional)*")
        st.divider()
        st.caption(
            "Detection uses explicit state/city names, address pincodes, and "
            "fuzzy spelling matches. Addresses that match nothing are labeled "
            "**Unknown** and listed separately so you can fix them by hand."
        )
    uploaded_file = st.file_uploader(
        "Upload Customer Report (Excel / CSV)",
        type=["xlsx", "csv"],
        help="File must include a 'Permanent Address' column."
    )
    if not uploaded_file:
        st.info("👋 Upload a customer report to get started. See the sidebar for the expected format.")
        st.stop()
    try:
        if uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file)
        else:
            df = pd.read_csv(uploaded_file)
    except Exception as e:
        st.error(f"❌ Couldn't read this file: {e}")
        st.stop()
    if df.empty:
        st.warning("⚠️ The uploaded file has no rows.")
        st.stop()
    if "Permanent Address" not in df.columns:
        st.error("❌ 'Permanent Address' column missing. Found columns: " + ", ".join(df.columns))
        st.stop()
    if "Temporary Address" not in df.columns:
        df["Temporary Address"] = ""
    if "Nearby Showroom" not in df.columns:
        df["Nearby Showroom"] = ""
    df["Final Address"] = df.apply(select_address, axis=1)
    progress_bar = st.progress(0, text="Detecting states and languages...")
    states, languages = [], []
    total_rows = len(df)
    chunk = max(1, total_rows // 100)  # update UI ~100 times max, not per row
    for i, addr in enumerate(df["Final Address"]):
        state, language = find_location(addr)
        states.append(state)
        languages.append(language)
        if i % chunk == 0 or i == total_rows - 1:
            progress_bar.progress((i + 1) / total_rows, text=f"Detecting states and languages... ({i+1:,}/{total_rows:,})")
    progress_bar.empty()
    df["State"] = states
    df["Language"] = languages
    df["Template"] = df["Language"].apply(lambda lang: "Tamil" if lang == "Tamil" else "English")
    st.success("✅ State, Language & Template added successfully")
    total_customers = len(df)
    tamil_customers = (df["Language"] == "Tamil").sum()
    unknown = (df["State"] == "Unknown").sum()
    other_languages = (df["Language"] != "Tamil").sum()
    non_tamil_identified = other_languages - unknown
    detected = total_customers - unknown
    detection_rate = (detected / total_customers * 100) if total_customers > 0 else 0
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📊 Total Customers", f"{total_customers:,}")
    c2.metric("🇮🇳 Tamil", f"{tamil_customers:,}", delta=f"{tamil_customers/total_customers*100:.1f}%")
    c3.metric("🌍 Other Languages", f"{non_tamil_identified:,}")
    c4.metric("❓ Unknown", f"{unknown:,}", delta=f"{unknown/total_customers*100:.1f}%", delta_color="inverse")
    st.progress(detection_rate / 100 if detection_rate > 0 else 0)
    if detection_rate < 50:
        st.error(f"⚠️ **Low Detection Rate: {detection_rate:.1f}%** — Only {detected:,} out of {total_customers:,} customers identified")
    else:
        st.success(f"✅ **Detection Rate: {detection_rate:.1f}%** — {detected:,} customers identified")
    st.divider()
    tab_overview, tab_results, tab_unknown, tab_download = st.tabs(
        ["📊 Overview", "📋 Full Results", "⚠️ Unknown Addresses", "📥 Downloads"]
    )
    with tab_overview:
        st.subheader("Customer Breakdown")
        summary_data = {
            "Category": ["Tamil Customers", "Other Languages (Identified)", "Unknown", "Total"],
            "Count": [tamil_customers, non_tamil_identified, unknown, total_customers],
            "Percentage": [
                f"{tamil_customers/total_customers*100:.1f}%",
                f"{non_tamil_identified/total_customers*100:.1f}%",
                f"{unknown/total_customers*100:.1f}%",
                "100%",
            ],
        }
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Language Distribution**")
            st.bar_chart(df["Language"].value_counts())
        with col_b:
            st.markdown("**Top 10 States**")
            st.bar_chart(df["State"].value_counts().head(10))
    with tab_results:
        st.subheader("Customer Result")

        fc1, fc2, fc3 = st.columns([2, 1, 1])
        with fc1:
            search_term = st.text_input("🔎 Search address", placeholder="Type part of an address to filter...")
        with fc2:
            state_filter = st.selectbox("Filter by state", ["All"] + sorted(df["State"].unique().tolist()))
        with fc3:
            language_filter = st.selectbox("Filter by language", ["All"] + sorted(df["Language"].unique().tolist()))

        filtered_df = df.copy()
        if search_term:
            filtered_df = filtered_df[filtered_df["Final Address"].str.contains(search_term, case=False, na=False)]
        if state_filter != "All":
            filtered_df = filtered_df[filtered_df["State"] == state_filter]
        if language_filter != "All":
            filtered_df = filtered_df[filtered_df["Language"] == language_filter]

        display_cols = ["Permanent Address", "Temporary Address", "Nearby Showroom",
                         "Final Address", "State", "Language", "Template"]
        display_cols = [col for col in display_cols if col in filtered_df.columns]

        st.caption(f"Showing {len(filtered_df):,} of {total_customers:,} customers")
        st.dataframe(filtered_df[display_cols], use_container_width=True, hide_index=True)
    with tab_unknown:
        unknown_df = df[df["State"] == "Unknown"]
        if len(unknown_df) > 0:
            st.warning(f"⚠️ **{len(unknown_df):,} customers with unknown state** ({len(unknown_df)/len(df)*100:.1f}% of total)")

            unique_unknown = unknown_df["Final Address"].value_counts()
            st.write("**Top 20 unknown addresses (most frequent first):**")
            st.dataframe(
                pd.DataFrame({"Address": unique_unknown.index[:20], "Count": unique_unknown.values[:20]}),
                use_container_width=True,
                hide_index=True,
            )

            csv_unknown = unknown_df[["Final Address"]].to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download Unknown Addresses (CSV)",
                csv_unknown,
                "unknown_addresses.csv",
                "text/csv",
            )
        else:
            st.success("🎉 No unknown addresses found! 100% detection rate!")
    with tab_download:
        st.subheader("Download the enriched report")
        col1, col2 = st.columns(2)

        with col1:
            excel_file = BytesIO()
            with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
            st.download_button(
                "📥 Download Excel (full report)",
                excel_file.getvalue(),
                "Customer_State_Language_Report.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col2:
            csv_file = df.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download CSV (full report)",
                csv_file,
                "Customer_State_Language_Report.csv",
                "text/csv",
                use_container_width=True,
            )

        with st.expander("📋 Sample Output Preview"):
            st.dataframe(df[["Final Address", "State", "Language", "Template"]].head(20), hide_index=True)
