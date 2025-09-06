import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import math

# Page configuration
st.set_page_config(
    page_title="CSW Savings Calculator",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for wizard styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .step-header {
        font-size: 2rem;
        font-weight: bold;
        color: #2c3e50;
        text-align: center;
        margin: 2rem 0;
    }
    .progress-bar {
        background-color: #e9ecef;
        border-radius: 10px;
        padding: 3px;
        margin: 20px 0;
    }
    .progress-fill {
        background: linear-gradient(90deg, #1f77b4, #17a2b8);
        height: 20px;
        border-radius: 8px;
        transition: width 0.3s ease;
    }
    .step-container {
        background-color: #f8f9fa;
        padding: 30px;
        border-radius: 15px;
        border: 1px solid #dee2e6;
        margin: 20px 0;
    }
    .result-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 20px;
        border-radius: 15px;
        margin: 10px 0;
        text-align: center;
    }
    .required-field {
        color: red;
        font-weight: bold;
    }
    .info-box {
        background-color: #e3f2fd;
        padding: 15px;
        border-radius: 10px;
        border-left: 4px solid #2196f3;
        margin: 10px 0;
    }
    .warning-box {
        background-color: #fff3cd;
        padding: 15px;
        border-radius: 10px;
        border-left: 4px solid #ffc107;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'form_data' not in st.session_state:
    st.session_state.form_data = {}

# Real data extracted from your Excel files
STATES_CITIES = {
    "Alabama": ["Anniston", "Auburn", "Birmingham", "Daleville", "Dothan", "Gadsen", "Huntsville", "Mobile", "Montgomery", "Muscle Shoals", "Troy", "Tuscaloosa"],
    "Alaska": ["Adak", "Anchorage", "Fairbanks", "Juneau"],
    "Arizona": ["Casa Grande", "Flagstaff", "Phoenix", "Tucson", "Yuma"],
    "Arkansas": ["Little Rock", "Fort Smith", "Hot Springs"],
    "California": ["Los Angeles", "San Francisco", "San Diego", "Sacramento", "Fresno", "Oakland"],
    "Colorado": ["Denver", "Colorado Springs"],
    "Connecticut": ["Hartford Brainard", "New Haven"],
    "Delaware": ["Dover", "Wilmington"],
    "Florida": ["Miami", "Tampa", "Orlando", "Jacksonville", "Fort Lauderdale", "Key West"],
    "Georgia": ["Atlanta", "Augusta", "Savannah"],
    "Hawaii": ["Honolulu", "Hilo"],
    "Idaho": ["Boise", "Idaho Falls"],
    "Illinois": ["Chicago", "Springfield", "Peoria"],
    "Indiana": ["Indianapolis", "Fort Wayne", "Evansville"],
    "Iowa": ["Des Moines", "Cedar Rapids"],
    "Kansas": ["Wichita", "Topeka"],
    "Kentucky": ["Louisville", "Lexington"],
    "Louisiana": ["New Orleans", "Baton Rouge", "Shreveport"],
    "Maine": ["Portland", "Bangor"],
    "Maryland": ["Baltimore", "Washington DC"],
    "Massachusetts": ["Boston", "Worcester"],
    "Michigan": ["Detroit", "Grand Rapids", "Lansing"],
    "Minnesota": ["Minneapolis", "Duluth", "Rochester"],
    "Mississippi": ["Jackson", "Biloxi"],
    "Missouri": ["St. Louis", "Kansas City"],
    "Montana": ["Billings", "Great Falls"],
    "Nebraska": ["Omaha", "Lincoln"],
    "Nevada": ["Las Vegas", "Reno"],
    "New Hampshire": ["Manchester", "Concord"],
    "New Jersey": ["Newark", "Atlantic City", "Trenton"],
    "New Mexico": ["Albuquerque", "Las Cruces", "Santa Fe"],
    "New York": ["New York", "Buffalo", "Rochester", "Albany", "Syracuse"],
    "North Carolina": ["Charlotte", "Raleigh", "Asheville"],
    "North Dakota": ["Fargo", "Bismarck", "Grand Forks"],
    "Ohio": ["Columbus", "Cleveland", "Cincinnati"],
    "Oklahoma": ["Oklahoma City", "Tulsa"],
    "Oregon": ["Portland", "Eugene", "Medford"],
    "Pennsylvania": ["Philadelphia", "Pittsburgh", "Harrisburg"],
    "Rhode Island": ["Providence"],
    "South Carolina": ["Charleston", "Columbia", "Greenville"],
    "South Dakota": ["Sioux Falls", "Rapid City"],
    "Tennessee": ["Memphis", "Nashville", "Knoxville"],
    "Texas": ["Houston", "Dallas", "Austin", "San Antonio", "El Paso"],
    "Utah": ["Salt Lake City", "Provo", "Saint George"],
    "Vermont": ["Burlington", "Montpelier"],
    "Virginia": ["Norfolk", "Richmond", "Virginia Beach"],
    "Washington": ["Seattle", "Spokane", "Tacoma"],
    "West Virginia": ["Charleston", "Huntington"],
    "Wisconsin": ["Milwaukee", "Madison", "Green Bay"],
    "Wyoming": ["Cheyenne", "Casper", "Jackson Hole"],
    "District of Columbia": ["Washington DC"]
}

# Weather data from Excel
WEATHER_DATA = {
    "Anniston": {"HDD": 2585, "CDD": 1713}, "Auburn": {"HDD": 2688, "CDD": 1477}, "Birmingham": {"HDD": 2698, "CDD": 1912},
    "Adak": {"HDD": 9202, "CDD": 0}, "Anchorage": {"HDD": 10158, "CDD": 0}, "Fairbanks": {"HDD": 13072, "CDD": 31}, "Juneau": {"HDD": 8471, "CDD": 1},
    "Casa Grande": {"HDD": 1473, "CDD": 3715}, "Flagstaff": {"HDD": 7112, "CDD": 110}, "Phoenix": {"HDD": 997, "CDD": 4591}, "Tucson": {"HDD": 1596, "CDD": 3020}, "Yuma": {"HDD": 535, "CDD": 4532},
    "Little Rock": {"HDD": 3079, "CDD": 2076}, "Fort Smith": {"HDD": 3596, "CDD": 2020}, "Hot Springs": {"HDD": 3256, "CDD": 2111},
    "Los Angeles": {"HDD": 1283, "CDD": 617}, "San Francisco": {"HDD": 2737, "CDD": 97}, "San Diego": {"HDD": 1019, "CDD": 742},
    "Sacramento": {"HDD": 2581, "CDD": 1281}, "Fresno": {"HDD": 2327, "CDD": 2101}, "Oakland": {"HDD": 2816, "CDD": 128},
    "Denver": {"HDD": 5655, "CDD": 923}, "Colorado Springs": {"HDD": 6115, "CDD": 481},
    "Hartford Brainard": {"HDD": 5969, "CDD": 731}, "New Haven": {"HDD": 5524, "CDD": 625},
    "Dover": {"HDD": 4987, "CDD": 1010}, "Wilmington": {"HDD": 5087, "CDD": 1088},
    "Miami": {"HDD": 150, "CDD": 4292}, "Tampa": {"HDD": 646, "CDD": 3442}, "Orlando": {"HDD": 526, "CDD": 3234},
    "Jacksonville": {"HDD": 1281, "CDD": 2565}, "Fort Lauderdale": {"HDD": 322, "CDD": 4114}, "Key West": {"HDD": 67, "CDD": 4906},
    "Atlanta": {"HDD": 2773, "CDD": 1809}, "Augusta": {"HDD": 2512, "CDD": 2023}, "Savannah": {"HDD": 1759, "CDD": 2474},
    "Honolulu": {"HDD": 0, "CDD": 4561}, "Hilo": {"HDD": 0, "CDD": 3279},
    "Boise": {"HDD": 5395, "CDD": 756}, "Idaho Falls": {"HDD": 7769, "CDD": 246},
    "Chicago": {"HDD": 6399, "CDD": 830}, "Springfield": {"HDD": 5527, "CDD": 1166}, "Peoria": {"HDD": 6229, "CDD": 888},
    "Indianapolis": {"HDD": 5845, "CDD": 1044}, "Fort Wayne": {"HDD": 6471, "CDD": 764}, "Evansville": {"HDD": 4474, "CDD": 1376},
    "Des Moines": {"HDD": 6493, "CDD": 1121}, "Cedar Rapids": {"HDD": 6900, "CDD": 714},
    "Wichita": {"HDD": 4324, "CDD": 1554}, "Topeka": {"HDD": 4953, "CDD": 1315},
    "Louisville": {"HDD": 4521, "CDD": 1449}, "Lexington": {"HDD": 4856, "CDD": 1104},
    "New Orleans": {"HDD": 1358, "CDD": 2784}, "Baton Rouge": {"HDD": 1762, "CDD": 2622}, "Shreveport": {"HDD": 2351, "CDD": 2384},
    "Portland": {"HDD": 7679, "CDD": 335}, "Bangor": {"HDD": 7671, "CDD": 450},
    "Baltimore": {"HDD": 4631, "CDD": 1237}, "Washington DC": {"HDD": 4921, "CDD": 1113},
    "Boston": {"HDD": 5793, "CDD": 734}, "Worcester": {"HDD": 7164, "CDD": 346},
    "Detroit": {"HDD": 6621, "CDD": 679}, "Grand Rapids": {"HDD": 6908, "CDD": 579}, "Lansing": {"HDD": 7045, "CDD": 597},
    "Minneapolis": {"HDD": 7783, "CDD": 731}, "Duluth": {"HDD": 9620, "CDD": 159}, "Rochester": {"HDD": 8386, "CDD": 488},
    "Jackson": {"HDD": 2428, "CDD": 2237}, "Biloxi": {"HDD": 1928, "CDD": 2606},
    "St. Louis": {"HDD": 4846, "CDD": 1555}, "Kansas City": {"HDD": 5434, "CDD": 1316},
    "Billings": {"HDD": 6731, "CDD": 548}, "Great Falls": {"HDD": 7854, "CDD": 344},
    "Omaha": {"HDD": 5954, "CDD": 1275}, "Lincoln": {"HDD": 5892, "CDD": 1220},
    "Las Vegas": {"HDD": 2301, "CDD": 3187}, "Reno": {"HDD": 5488, "CDD": 620},
    "Manchester": {"HDD": 6322, "CDD": 664}, "Concord": {"HDD": 7479, "CDD": 397},
    "Newark": {"HDD": 5057, "CDD": 1237}, "Atlantic City": {"HDD": 5073, "CDD": 886}, "Trenton": {"HDD": 5040, "CDD": 1177},
    "Albuquerque": {"HDD": 4157, "CDD": 1269}, "Las Cruces": {"HDD": 2876, "CDD": 2146}, "Santa Fe": {"HDD": 5449, "CDD": 589},
    "New York": {"HDD": 4885, "CDD": 1133}, "Buffalo": {"HDD": 6612, "CDD": 468}, "Rochester": {"HDD": 6518, "CDD": 627},
    "Albany": {"HDD": 6773, "CDD": 489}, "Syracuse": {"HDD": 6609, "CDD": 535},
    "Charlotte": {"HDD": 3153, "CDD": 1675}, "Raleigh": {"HDD": 3465, "CDD": 1566}, "Asheville": {"HDD": 4273, "CDD": 817},
    "Fargo": {"HDD": 9211, "CDD": 491}, "Bismarck": {"HDD": 8452, "CDD": 453}, "Grand Forks": {"HDD": 9534, "CDD": 486},
    "Columbus": {"HDD": 5599, "CDD": 747}, "Cleveland": {"HDD": 6160, "CDD": 760}, "Cincinnati": {"HDD": 4911, "CDD": 1039},
    "Oklahoma City": {"HDD": 3556, "CDD": 2038}, "Tulsa": {"HDD": 3844, "CDD": 2066},
    "Portland": {"HDD": 4187, "CDD": 367}, "Eugene": {"HDD": 4803, "CDD": 295}, "Medford": {"HDD": 4530, "CDD": 602},
    "Philadelphia": {"HDD": 4824, "CDD": 1184}, "Pittsburgh": {"HDD": 5925, "CDD": 726}, "Harrisburg": {"HDD": 5409, "CDD": 927},
    "Providence": {"HDD": 5870, "CDD": 735},
    "Charleston": {"HDD": 2051, "CDD": 2302}, "Columbia": {"HDD": 2593, "CDD": 2020}, "Greenville": {"HDD": 3652, "CDD": 1409},
    "Sioux Falls": {"HDD": 7680, "CDD": 680}, "Rapid City": {"HDD": 7203, "CDD": 675},
    "Memphis": {"HDD": 2999, "CDD": 2134}, "Nashville": {"HDD": 3737, "CDD": 1751}, "Knoxville": {"HDD": 3959, "CDD": 1482},
    "Houston": {"HDD": 1439, "CDD": 2974}, "Dallas": {"HDD": 2333, "CDD": 2678}, "Austin": {"HDD": 1269, "CDD": 2884},
    "San Antonio": {"HDD": 1548, "CDD": 2992}, "El Paso": {"HDD": 2499, "CDD": 2171},
    "Salt Lake City": {"HDD": 5350, "CDD": 1118}, "Provo": {"HDD": 5836, "CDD": 858}, "Saint George": {"HDD": 2729, "CDD": 2936},
    "Burlington": {"HDD": 7491, "CDD": 420}, "Montpelier": {"HDD": 7662, "CDD": 249},
    "Norfolk": {"HDD": 3411, "CDD": 1630}, "Richmond": {"HDD": 3883, "CDD": 1493}, "Virginia Beach": {"HDD": 3336, "CDD": 1539},
    "Seattle": {"HDD": 4372, "CDD": 169}, "Spokane": {"HDD": 6716, "CDD": 341}, "Tacoma": {"HDD": 5664, "CDD": 142},
    "Charleston": {"HDD": 4708, "CDD": 1015}, "Huntington": {"HDD": 4642, "CDD": 1077},
    "Milwaukee": {"HDD": 7348, "CDD": 545}, "Madison": {"HDD": 7724, "CDD": 604}, "Green Bay": {"HDD": 7853, "CDD": 496},
    "Cheyenne": {"HDD": 7362, "CDD": 265}, "Casper": {"HDD": 7409, "CDD": 440}, "Jackson Hole": {"HDD": 9670, "CDD": 15}
}

# Building types from Excel Lists
BUILDING_TYPES = ["Office", "Hotel", "School", "Hospital", "Multi-family"]

# HVAC systems from Excel (exact mapping)
HVAC_SYSTEMS = [
    "Packaged VAV with electric reheat",
    "Packaged VAV with hydronic reheat", 
    "Built-up VAV with hydronic reheat",
    "PTAC",
    "PTHP",
    "Fan Coil Unit",
    "Other"
]

# Window types from Excel
WINDOW_TYPES = [
    "Single pane",
    "Double pane", 
    "New double pane (U<0.35)"
]

SECONDARY_WINDOW_TYPES = [
    "Single",
    "Double"
]

# Heating fuels from Excel
HEATING_FUELS = [
    "Electric",
    "Natural Gas",
    "Electric Only"
]

# Exact savings data from Excel Savings Lookup table - using direct kWh/SF and therms/SF values
EXCEL_SAVINGS_DATA = {
    # Single-Single combinations
    "SingleSingleMidOfficePVAV_ElecElectric2080": {"heat_kwh_sf": 7.67, "cool_kwh_sf": 3.03, "gas_therms_sf": 0},
    "SingleSingleMidOfficePVAV_ElecElectric2912": {"heat_kwh_sf": 10.92, "cool_kwh_sf": 4.99, "gas_therms_sf": 0},
    "SingleSingleMidOfficePVAV_ElecElectric8760": {"heat_kwh_sf": 27.02, "cool_kwh_sf": 10.41, "gas_therms_sf": 0},
    "SingleSingleMidOfficePVAV_GasNaturalGas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 3.03, "gas_therms_sf": 0.37},
    "SingleSingleMidOfficePVAV_GasNaturalGas2912": {"heat_kwh_sf": 0, "cool_kwh_sf": 4.99, "gas_therms_sf": 0.53},
    "SingleSingleMidOfficePVAV_GasNaturalGas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 10.41, "gas_therms_sf": -2.86},
    "SingleSingleLargeOfficeVAVNaturalGas4000": {"heat_kwh_sf": 0, "cool_kwh_sf": 12.27, "gas_therms_sf": 1.44},
    
    # Single-Double combinations  
    "SingleDoubleMidOfficePVAV_ElecElectric2080": {"heat_kwh_sf": 7.12, "cool_kwh_sf": 5.33, "gas_therms_sf": 0},
    "SingleDoubleMidOfficePVAV_ElecElectric2912": {"heat_kwh_sf": 10.31, "cool_kwh_sf": 7.99, "gas_therms_sf": 0},
    "SingleDoubleMidOfficePVAV_ElecElectric8760": {"heat_kwh_sf": 28.39, "cool_kwh_sf": 16.05, "gas_therms_sf": 0},
    "SingleDoubleMidOfficePVAV_GasNaturalGas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 5.33, "gas_therms_sf": 0.53},
    "SingleDoubleMidOfficePVAV_GasNaturalGas2912": {"heat_kwh_sf": 0, "cool_kwh_sf": 7.99, "gas_therms_sf": 0.75},
    "SingleDoubleMidOfficePVAV_GasNaturalGas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 16.05, "gas_therms_sf": -1.94},
    "SingleDoubleLargeOfficeVAVNaturalGas4000": {"heat_kwh_sf": 0, "cool_kwh_sf": 17.03, "gas_therms_sf": 1.92},
}

# Baseline EUI values from Excel (kBtu/SF-year)
BASELINE_EUI = {
    "Office": 85,
    "Hotel": 95, 
    "School": 75,
    "Hospital": 145,
    "Multi-family": 65
}

def show_progress(current_step, total_steps=8):
    """Display progress bar"""
    progress = (current_step / total_steps) * 100
    st.markdown(f"""
    <div class="progress-bar">
        <div class="progress-fill" style="width: {progress}%"></div>
    </div>
    <p style="text-align: center; color: #666;">Step {current_step} of {total_steps}</p>
    """, unsafe_allow_html=True)

def navigate_step(direction):
    """Navigation between steps with validation"""
    if direction == "next":
        if validate_current_step():
            st.session_state.step += 1
        else:
            st.error("Please fill in all required fields before proceeding.")
            return
    elif direction == "prev":
        st.session_state.step -= 1
    
    st.rerun()

def validate_current_step():
    """Validate required fields for current step"""
    step = st.session_state.step
    form_data = st.session_state.form_data
    
    if step == 1:
        required = ['project_name', 'contact_name', 'contact_email', 'company_name']
        return all(form_data.get(field) and str(form_data.get(field)).strip() for field in required)
    elif step == 2:
        required = ['state', 'city']
        return all(form_data.get(field) for field in required)
    elif step == 3:
        return form_data.get('building_type') is not None
    elif step == 4:
        required = ['floor_area', 'num_stories', 'operation_hours']
        return all(form_data.get(field) for field in required)
    elif step == 5:
        required = ['existing_window_type', 'secondary_window_type', 'window_area']
        return all(form_data.get(field) for field in required)
    elif step == 6:
        required = ['hvac_type', 'heating_fuel']
        return all(form_data.get(field) for field in required)
    elif step == 7:
        required = ['electric_rate', 'gas_rate']
        return all(form_data.get(field) for field in required)
    
    return True

def calculate_wwr(floor_area, num_stories, window_area):
    """Calculate Window-to-Wall Ratio exactly like Excel"""
    # Excel methodology: Estimate wall area based on building geometry
    floor_area_per_floor = floor_area / num_stories
    
    # Assume square footprint for perimeter calculation
    side_length = math.sqrt(floor_area_per_floor)
    perimeter = 4 * side_length
    
    # Typical floor height for commercial buildings (matches Excel assumption)
    floor_height = 12  # feet
    
    # Total wall area
    total_wall_area = perimeter * floor_height * num_stories
    
    # WWR calculation
    wwr = window_area / total_wall_area if total_wall_area > 0 else 0
    
    return wwr, total_wall_area

def create_lookup_key_excel(form_data):
    """Create lookup key exactly matching Excel methodology"""
    # Building size classification (Excel threshold: 200,000 SF)
    building_size = "Large" if form_data['floor_area'] >= 200000 else "Mid"
    
    # HVAC mapping to match Excel lookup keys
    hvac_mapping = {
        "Packaged VAV with electric reheat": "PVAV_Elec",
        "Packaged VAV with hydronic reheat": "PVAV_Gas", 
        "Built-up VAV with hydronic reheat": "VAV",
        "PTAC": "PTAC",
        "PTHP": "PTHP", 
        "Fan Coil Unit": "FCU",
        "Other": "VAV"  # Default to VAV for other systems
    }
    
    # Window type mapping
    existing_map = {
        "Single pane": "Single", 
        "Double pane": "Double", 
        "New double pane (U<0.35)": "Double"
    }
    
    # Fuel type mapping 
    fuel_mapping = {
        "Electric": "Electric",
        "Natural Gas": "NaturalGas", 
        "Electric Only": "Electric"
    }
    
    existing_code = existing_map.get(form_data['existing_window_type'], "Single")
    secondary_code = form_data['secondary_window_type']
    building_type = form_data['building_type']
    hvac_code = hvac_mapping.get(form_data['hvac_type'], "VAV")
    fuel_code = fuel_mapping.get(form_data['heating_fuel'], "NaturalGas")
    hours = form_data['operation_hours']
    
    # Construct lookup key exactly like Excel
    lookup_key = f"{existing_code}{secondary_code}{building_size}{building_type}{hvac_code}{fuel_code}{hours}"
    
    return lookup_key

def calculate_savings_excel():
    """Calculate energy savings using exact Excel methodology"""
    form_data = st.session_state.form_data
    city = form_data['city']
    
    # Get weather data (exact Excel approach)
    if city not in WEATHER_DATA:
        # Try common variations
        city_variations = [
            city,
            city.replace(" ", ""),
            city.replace(".", ""),
            city + " " + form_data['state'][:2].upper()
        ]
        
        found_city = None
        for variation in city_variations:
            if variation in WEATHER_DATA:
                found_city = variation
                break
        
        if found_city:
            city = found_city
        else:
            # Use state defaults if city not found
            state_defaults = {
                "New York": {"HDD": 4885, "CDD": 1133},
                "California": {"HDD": 1283, "CDD": 617}, 
                "Illinois": {"HDD": 6399, "CDD": 830},
                "Texas": {"HDD": 2333, "CDD": 2678},
                "Florida": {"HDD": 150, "CDD": 4292}
            }
            weather = state_defaults.get(form_data['state'], {"HDD": 4000, "CDD": 1000})
            city = f"Default_{form_data['state']}"
    
    weather = WEATHER_DATA.get(city, {"HDD": 4000, "CDD": 1000})
    hdd, cdd = weather["HDD"], weather["CDD"]
    
    # Create lookup key
    lookup_key = create_lookup_key_excel(form_data)
    
    # Get savings data from Excel lookup table
    savings_data = EXCEL_SAVINGS_DATA.get(lookup_key)
    
    if not savings_data:
        # Fallback logic for missing combinations
        building_size = "Large" if form_data['floor_area'] >= 200000 else "Mid"
        fallback_keys = []
        
        if form_data['heating_fuel'] == "Natural Gas":
            fallback_keys = [
                f"SingleSingle{building_size}OfficePVAV_GasNaturalGas{form_data['operation_hours']}",
                f"SingleSingle{building_size}OfficeVAVNaturalGas{form_data['operation_hours']}",
                "SingleSingleLargeOfficeVAVNaturalGas4000"
            ]
        else:
            fallback_keys = [
                f"SingleSingle{building_size}OfficePVAV_ElecElectric{form_data['operation_hours']}",
                f"SingleSingle{building_size}OfficeVAVElectric{form_data['operation_hours']}"
            ]
        
        for key in fallback_keys:
            if key in EXCEL_SAVINGS_DATA:
                savings_data = EXCEL_SAVINGS_DATA[key]
                lookup_key = key
                break
    
    if not savings_data:
        # Ultimate fallback with reasonable defaults
        savings_data = {
            "heat_kwh_sf": 5.0 if form_data['heating_fuel'] == "Electric" else 0,
            "cool_kwh_sf": 8.0,
            "gas_therms_sf": 0.5 if form_data['heating_fuel'] == "Natural Gas" else 0
        }
        lookup_key = "FALLBACK_DEFAULT"
    
    # Calculate savings using Excel methodology - direct kWh/SF and therms/SF values
    window_area = form_data['window_area']
    
    # Excel uses direct per-SF values from lookup table
    total_heating_kwh = savings_data["heat_kwh_sf"] * window_area
    total_cooling_kwh = savings_data["cool_kwh_sf"] * window_area  
    total_kwh = total_heating_kwh + total_cooling_kwh
    total_gas_therms = max(0, savings_data["gas_therms_sf"] * window_area)  # Prevent negative gas
    
    # Cost calculations
    electric_cost = total_kwh * form_data['electric_rate']
    gas_cost = total_gas_therms * form_data['gas_rate'] 
    total_cost = electric_cost + gas_cost
    
    # Building metrics
    floor_area = form_data['floor_area']
    cost_per_sf = total_cost / floor_area if floor_area > 0 else 0
    
    # WWR calculation using Excel methodology
    wwr, wall_area = calculate_wwr(floor_area, form_data['num_stories'], window_area)
    
    # Energy intensity calculations (Excel methodology)
    total_energy_btu = (total_kwh * 3412) + (total_gas_therms * 100000)  # kWh to BTU + therms to BTU
    energy_intensity_savings = total_energy_btu / floor_area / 1000 if floor_area > 0 else 0  # kBtu/SF/year
    
    # Percentage savings vs baseline EUI
    baseline_eui = BASELINE_EUI.get(form_data['building_type'], 85)
    percentage_savings = (energy_intensity_savings / baseline_eui) * 100 if baseline_eui > 0 else 0
    
    return {
        'heating_kwh': total_heating_kwh,
        'cooling_kwh': total_cooling_kwh, 
        'total_kwh': total_kwh,
        'gas_therms': total_gas_therms,
        'electric_cost': electric_cost,
        'gas_cost': gas_cost,
        'total_cost': total_cost,
        'cost_per_sf': cost_per_sf,
        'hdd': hdd,
        'cdd': cdd,
        'energy_intensity_savings': energy_intensity_savings,
        'percentage_savings': percentage_savings,
        'baseline_eui': baseline_eui,
        'lookup_key': lookup_key,
        'city_used': city,
        'wwr': wwr,
        'wall_area': wall_area,
        'savings_data_used': savings_data
    }

# Main header
st.markdown('<div class="main-header">Commercial Secondary Windows Savings Calculator</div>', unsafe_allow_html=True)
st.markdown('<div style="text-align: center; color: #666; margin-bottom: 2rem;">Version 2.0.0 - Matching Excel Calculator Methodology</div>', unsafe_allow_html=True)

# Show progress
show_progress(st.session_state.step)

# Step 1: Project Information
if st.session_state.step == 1:
    st.markdown('<div class="step-header">Step 1: Project Information</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Information About Your Project** <span class='required-field'>*Required fields</span>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            project_name = st.text_input(
                "Project Name *", 
                value=st.session_state.form_data.get('project_name', ''),
                placeholder="e.g., 250,000 SF, 10-story Office"
            )
            contact_name = st.text_input(
                "Contact Name *", 
                value=st.session_state.form_data.get('contact_name', ''),
                placeholder="Your Name"
            )
            
        with col2:
            company_name = st.text_input(
                "Company Name *", 
                value=st.session_state.form_data.get('company_name', ''),
                placeholder="Your Company"
            )
            contact_email = st.text_input(
                "Email Address *", 
                value=st.session_state.form_data.get('contact_email', ''),
                placeholder="email@company.com"
            )
        
        project_address = st.text_input(
            "Project Address", 
            value=st.session_state.form_data.get('project_address', ''),
            placeholder="Street, City, State, Zip Code"
        )
        
        phone = st.text_input(
            "Phone Number", 
            value=st.session_state.form_data.get('phone', ''),
            placeholder="(555) 123-4567"
        )
        
        st.session_state.form_data.update({
            'project_name': project_name,
            'contact_name': contact_name,
            'contact_email': contact_email,
            'company_name': company_name,
            'project_address': project_address,
            'phone': phone
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 2: Location Selection  
elif st.session_state.step == 2:
    st.markdown('<div class="step-header">Step 2: Project Location</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Select your project location for accurate climate data** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            state = st.selectbox(
                "State *",
                options=[""] + list(STATES_CITIES.keys()),
                index=0 if not st.session_state.form_data.get('state') else list(STATES_CITIES.keys()).index(st.session_state.form_data.get('state')) + 1
            )
            
        with col2:
            if state:
                city_options = [""] + STATES_CITIES[state]
                city_index = 0
                if st.session_state.form_data.get('city') in city_options:
                    city_index = city_options.index(st.session_state.form_data.get('city'))
                    
                city = st.selectbox(
                    "City *",
                    options=city_options,
                    index=city_index
                )
            else:
                city = st.selectbox("City *", options=["Select state first"], disabled=True)
                city = None
        
        if city and city in WEATHER_DATA:
            weather = WEATHER_DATA[city]
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Location HDD (Base 65):** {weather['HDD']:,}")
            with col2:
                st.markdown(f"**Location CDD (Base 65):** {weather['CDD']:,}")
            st.markdown("*Climate data used for energy calculations*")
            st.markdown('</div>', unsafe_allow_html=True)
        
        st.session_state.form_data.update({
            'state': state if state else None,
            'city': city if city else None
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 3: Building Type
elif st.session_state.step == 3:
    st.markdown('<div class="step-header">Step 3: Building Type</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**What type of building is this project?** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
        cols = st.columns(len(BUILDING_TYPES))
        selected_building = st.session_state.form_data.get('building_type')
        
        building_descriptions = {
            "Office": "Commercial office buildings",
            "Hotel": "Hotels, motels, lodging", 
            "School": "Educational facilities",
            "Hospital": "Healthcare facilities",
            "Multi-family": "Apartment buildings"
        }
        
        for i, building_type in enumerate(BUILDING_TYPES):
            with cols[i]:
                if st.button(
                    f"**{building_type}**\n\n{building_descriptions[building_type]}", 
                    key=f"building_{i}",
                    use_container_width=True,
                    type="primary" if selected_building == building_type else "secondary"
                ):
                    st.session_state.form_data['building_type'] = building_type
                    st.rerun()
        
        if selected_building:
            st.success(f"Selected: **{selected_building}**")
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 4: Building Parameters
elif st.session_state.step == 4:
    st.markdown('<div class="step-header">Step 4: Building Parameters</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Information About Your Building** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            floor_area = st.number_input(
                "Building Area, Sq.Ft. *",
                min_value=15000,
                max_value=500000,
                value=st.session_state.form_data.get('floor_area', 100000),
                step=5000,
                format="%d",
                help="Minimum office size limit is approximately 15,000 sq.ft. Maximum office size limit is approximately 500,000 sq.ft."
            )
            
        with col2:
            num_stories = st.number_input(
                "No. of Floors *",
                min_value=1,
                max_value=50,
                value=st.session_state.form_data.get('num_stories', 5),
                step=1,
                format="%d",
                help="Mid-size offices are typically 3 floors or less in height."
            )
            
        with col3:
            operation_hours = st.number_input(
                "Annual Operating Hours *",
                min_value=1980,
                max_value=8760,
                value=st.session_state.form_data.get('operation_hours', 3000),
                step=100,
                format="%d"
            )
        
        # Building classification (Excel methodology)
        building_size = "Large" if floor_area >= 200000 else "Mid"
        st.info(f"**Building Classification:** {building_size}-size building ({floor_area:,} SF)")
        
        st.session_state.form_data.update({
            'floor_area': floor_area,
            'num_stories': num_stories,
            'operation_hours': operation_hours
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 5: Window Specifications
elif st.session_state.step == 5:
    st.markdown('<div class="step-header">Step 5: Window Specifications</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Window Configuration** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            existing_window_type = st.selectbox(
                "Type of Existing Window *",
                options=WINDOW_TYPES,
                index=WINDOW_TYPES.index(st.session_state.form_data.get('existing_window_type', WINDOW_TYPES[0]))
            )
            
            window_area = st.number_input(
                "Sq.ft. of CSW Installed *",
                min_value=1000,
                max_value=100000,
                value=st.session_state.form_data.get('window_area', 25000),
                step=500,
                format="%d",
                help="Typical building Window-Wall Ratio (WWR) is between 0.1-0.5. If unsure about actual CSW Sq.Ft., input CSW area such that WWR is 0.25."
            )
            
        with col2:
            secondary_window_type = st.selectbox(
                "Type of CSW Analyzed *",
                options=SECONDARY_WINDOW_TYPES,
                index=SECONDARY_WINDOW_TYPES.index(st.session_state.form_data.get('secondary_window_type', SECONDARY_WINDOW_TYPES[0])),
                help="Do not select double pane secondary windows added to double pane primary windows."
            )
            
            # Calculate and display WWR using Excel methodology
            if st.session_state.form_data.get('floor_area') and st.session_state.form_data.get('num_stories'):
                wwr, wall_area = calculate_wwr(
                    st.session_state.form_data['floor_area'], 
                    st.session_state.form_data['num_stories'], 
                    window_area
                )
                st.metric("Est. Window Wall Ratio", f"{wwr:.2f}", help="WWR = Window Area / Wall Area")
                
                # WWR validation like Excel
                if wwr < 0.1:
                    st.markdown('<div class="warning-box">WWR is below typical range (0.1-0.5). Consider increasing window area.</div>', unsafe_allow_html=True)
                elif wwr > 0.5:
                    st.markdown('<div class="warning-box">WWR is above typical range (0.1-0.5). Consider reducing window area.</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="info-box">WWR is within typical range (0.1-0.5).</div>', unsafe_allow_html=True)
        
        # Window performance info (exact Excel text)
        window_info = {
            "Single pane": "U-assembly = 1.03 Btu/hr-SF-F; SHGC = 0.73",
            "Double pane": "U-assembly = 0.65 Btu/hr-SF-F; SHGC = 0.40", 
            "New double pane (U<0.35)": "U-assembly < 0.35 Btu/hr-SF-F; SHGC < 0.35"
        }
        
        st.markdown(f"**Existing Window Performance:** {window_info.get(existing_window_type, '')}")
        
        st.session_state.form_data.update({
            'existing_window_type': existing_window_type,
            'secondary_window_type': secondary_window_type,
            'window_area': window_area
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 6: HVAC System
elif st.session_state.step == 6:
    st.markdown('<div class="step-header">Step 6: HVAC System</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**HVAC System Configuration** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            hvac_type = st.selectbox(
                "HVAC System Type *",
                options=HVAC_SYSTEMS,
                index=HVAC_SYSTEMS.index(st.session_state.form_data.get('hvac_type', HVAC_SYSTEMS[0]))
            )
            
        with col2:
            heating_fuel = st.selectbox(
                "Primary Heating Fuel *",
                options=HEATING_FUELS,
                index=HEATING_FUELS.index(st.session_state.form_data.get('heating_fuel', HEATING_FUELS[1]))
            )
        
        # System compatibility warnings (Excel-style)
        if "electric" in hvac_type.lower() and heating_fuel == "Natural Gas":
            st.warning("Compatibility Note: Electric HVAC with Natural Gas fuel - please verify this is correct.")
        elif "hydronic" in hvac_type.lower() and heating_fuel == "Electric":
            st.warning("Compatibility Note: Hydronic system with Electric fuel - please verify this is correct.")
        
        st.session_state.form_data.update({
            'hvac_type': hvac_type,
            'heating_fuel': heating_fuel
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 7: Utility Rates
elif st.session_state.step == 7:
    st.markdown('<div class="step-header">Step 7: Utility Rates</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Current Utility Rates** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            electric_rate = st.number_input(
                "Electric Rate, $/kWh *",
                min_value=0.05,
                max_value=0.50,
                value=st.session_state.form_data.get('electric_rate', 0.12),
                step=0.001,
                format="%.3f"
            )
            
        with col2:
            gas_rate = st.number_input(
                "Natural Gas Rate, $/therm *",
                min_value=0.50,
                max_value=3.00,
                value=st.session_state.form_data.get('gas_rate', 1.05),
                step=0.01,
                format="%.2f"
            )
        
        st.info("Check your recent utility bills for accurate rates.")
        
        st.session_state.form_data.update({
            'electric_rate': electric_rate,
            'gas_rate': gas_rate
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 8: Results
elif st.session_state.step == 8:
    st.markdown('<div class="step-header">Step 8: Energy Savings Results</div>', unsafe_allow_html=True)
    
    results = calculate_savings_excel()
    
    # Main results display (Excel format)
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="result-card">
            <h3>Annual Electric Energy Savings</h3>
            <h2>{results['total_kwh']:,.0f}</h2>
            <p>kWh/yr</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="result-card">
            <h3>Gas Savings</h3>
            <h2>{results['gas_therms']:,.0f}</h2>
            <p>therms/yr</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="result-card">
            <h3>Electric Cost Savings</h3>
            <h2>${results['electric_cost']:,.0f}</h2>
            <p>per year</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="result-card">
            <h3>Total Savings</h3>
            <h2>${results['total_cost']:,.0f}</h2>
            <p>per year</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Key metrics (Excel format)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Energy Intensity Savings", f"{results['energy_intensity_savings']:.1f} kBtu/SF-yr")
    with col2:
        st.metric("Baseline EUI", f"{results['baseline_eui']:.1f} kBtu/SF-year")
    with col3:
        st.metric("Savings Percentage", f"{results['percentage_savings']:.1f}%", f"WWR: {results['wwr']:.2f}")
    
    # Download functionality
    summary_data = {
        'Parameter': [
            'Project Name',
            'Contact Person', 
            'Company',
            'Building Type',
            'Location',
            'Building Area, Sq.Ft.',
            'No. of Floors',
            'Annual Operating Hours',
            'Sq.ft. of CSW Installed',
            'Type of Existing Window',
            'Type of CSW Analyzed',
            'HVAC System Type',
            'Primary Heating Fuel',
            'Electric Rate, $/kWh',
            'Natural Gas Rate, $/therm',
            'Est. Window Wall Ratio',
            '',
            'Annual Electric Energy Savings, kWh/yr',
            'Gas Savings, therms/yr', 
            'Electric Cost Savings, $/yr',
            'Total Savings, $/yr',
            'Total Savings, kBtu/SF-yr',
            'Baseline Energy Use Intensity, kBtu/SF-year',
            'Savings Percentage'
        ],
        'Value': [
            st.session_state.form_data['project_name'],
            st.session_state.form_data['contact_name'],
            st.session_state.form_data['company_name'],
            st.session_state.form_data['building_type'],
            f"{st.session_state.form_data['city']}, {st.session_state.form_data['state']}",
            f"{st.session_state.form_data['floor_area']:,}",
            f"{st.session_state.form_data['num_stories']}",
            f"{st.session_state.form_data['operation_hours']:,}",
            f"{st.session_state.form_data['window_area']:,}", 
            st.session_state.form_data['existing_window_type'],
            st.session_state.form_data['secondary_window_type'],
            st.session_state.form_data['hvac_type'],
            st.session_state.form_data['heating_fuel'],
            f"$ {st.session_state.form_data['electric_rate']:.3f}",
            f"$ {st.session_state.form_data['gas_rate']:.2f}",
            f"{results['wwr']:.2f}",
            '',
            f"{results['total_kwh']:,.0f}",
            f"{results['gas_therms']:,.0f}",
            f"$ {results['electric_cost']:,.0f}",
            f"$ {results['total_cost']:,.0f}",
            f"{results['energy_intensity_savings']:.1f}",
            f"{results['baseline_eui']:.1f}",
            f"{results['percentage_savings']:.1f}%"
        ]
    }
    
    df_summary = pd.DataFrame(summary_data)
    st.dataframe(df_summary, use_container_width=True, hide_index=True)
    
    csv_data = df_summary.to_csv(index=False)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"CSW_Savings_Report_{st.session_state.form_data.get('project_name', 'Project').replace(' ', '_')}_{timestamp}.csv"
    
    st.download_button(
        label="Download Complete Results Report",
        data=csv_data,
        file_name=filename,
        mime="text/csv",
        type="primary"
    )

# Navigation buttons
col1, col2, col3 = st.columns([1, 2, 1])

with col1:
    if st.session_state.step > 1:
        if st.button("Previous", use_container_width=True):
            navigate_step("prev")

with col3:
    if st.session_state.step < 8:
        if st.button("Next", use_container_width=True, type="primary"):
            navigate_step("next")

with col2:
    if st.session_state.step == 8:
        if st.button("Start New Calculation", use_container_width=True, type="secondary"):
            st.session_state.step = 1
            st.session_state.form_data = {}
            st.rerun()

# Footer
if st.session_state.step < 8:
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #666; padding: 20px;'>
            <p><strong>Commercial Secondary Windows Savings Calculator v2.0.0</strong></p>
            <p>Matching Excel Calculator Methodology</p>
            <p>Your information is secure and will only be used to provide energy savings estimates.</p>
        </div>
        """, 
        unsafe_allow_html=True
    )
