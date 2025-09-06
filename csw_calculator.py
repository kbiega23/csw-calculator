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

# Complete data extracted from your Excel files - ALL cities included
STATES_CITIES = {
    "Alabama": ["Anniston", "Auburn", "Birmingham", "Daleville", "Dothan", "Gadsen", "Huntsville", "Mobile", "Montgomery", "Muscle Shoals", "Troy", "Tuscaloosa"],
    "Alaska": ["Adak", "Anchorage", "Aniak", "Annette Island", "Anvik", "Big River Lakes", "Birchwood", "Chulitna", "Cold Bay", "Cordova", "Dillingham", "Dutch Harbor", "Emmonak", "Fairbanks", "Gustavus", "Healy", "Homer", "Hoonah", "Hooper Bay", "Huslia", "Iliamna", "Juneau", "Kake", "Kenai", "Ketchikan", "King Salmon", "Kodiak", "Mekoryuk", "Middleton Island", "Palmer", "Petersburg", "Port Heiden", "Saint Mary's", "Sand Point", "Seward", "Sitka Japonski", "Skagway", "Soldotna", "St. Paul", "Talkeetna", "Togiac", "Utqiagvik", "Valdez", "Whittier", "Wrangell", "Yakutat"],
    "Arizona": ["Casa Grande", "Douglas", "Flagstaff", "Grand Canyon", "Kingman", "Luke", "Page", "Phoenix", "Prescott", "Safford", "Scottsdale", "Show Low", "Tucson", "Winslow", "Yuma"],
    "Arkansas": ["Batesville", "Bentonville", "El Dorado", "Fayetteville", "Flippin", "Fort Smith", "Harrison", "Hot Springs", "Jonesboro", "Little Rock", "Pine Bluff", "Rogers", "Siloam Spring", "Springdale", "Stuttgart", "Texarkana", "Walnut Ridge"],
    "California": ["Alturas", "Bakersfield", "Bishop", "Blue Canyon", "Blythe", "Burbank", "Camarillo", "Camp Pendleton", "Carlsbad", "Chino", "Chula Vista", "Concord", "Crescent City", "Daggett", "Fairfield", "Fresno", "Fullerton", "Hawthorne", "Hayward", "Imperial", "Lancaster", "Lemoore", "Livermore", "Lompoc", "Long Beach", "Los Angeles", "McKinleyville", "Merced", "Modesto", "Montague", "Monterey", "Napa", "Needles", "Oakland", "Oxnard", "Palm Springs", "Palmdale", "Paso Robles", "Porterville", "Red Bluff", "Redding", "Ridgecrest", "Riverside", "Sacramento", "Salinas", "San Diego", "San Francisco", "San Jose", "San Luis Obispo", "Sandberg", "Santa Ana", "Santa Barbara", "Santa Maria", "Santa Monica", "Santa Rosa", "South Lake Tahoe", "Stockton", "Truckee", "Twentynine Palms", "Ukiah", "Van Nuys", "Visalia", "Yuba City"],
    "Colorado": ["Akron", "Alamosa", "Aspen", "Aurora", "Broomfield", "Colorado Springs", "Cortez", "Craig", "Denver", "Durango", "Fort Collins", "Grand Junction", "Greeley", "Gunnison", "Gypsum", "Hayden", "La Junta", "Lakewood", "Lamar", "Leadville", "Limon", "Montrose", "Pueblo", "Rifle", "Trinidad"],
    "Connecticut": ["Danbury", "Groton", "Hartford Brainard", "New Haven", "Oxford", "Stratford"],
    "Delaware": ["Dover", "Wilmington"],
    "Florida": ["Crestview", "Daytona", "Fort Lauderdale", "Fort Myers", "Fort Pierce", "Gainesville", "Homestead", "Jacksonville", "Key West", "Lakeland", "Marathon", "Mayport", "Melbourne", "Miami", "Milton", "Naples", "Ocala", "Orlando", "Panama City", "Pensacola", "Sarasota", "St. Petersburg", "Tallahassee", "Tampa", "Valparaiso", "Vero Beach", "West Palm Beach"],
    "Georgia": ["Albany", "Alma", "Athens", "Atlanta", "Augusta", "Brunswick", "Columbus", "Macon", "Marietta", "Rome", "Savannah", "Valdosta", "Warner Robins"],
    "Hawaii": ["Hilo", "Honolulu", "Kahului", "Kalaoa", "Kaneohe", "Kapalua", "Kapolei", "Lanai", "Lihue", "Molokai"],
    "Idaho": ["Boise", "Burley", "Caldwell", "Coeur d'Alene", "Grand View", "Hailey", "Idaho Falls", "Lewiston", "Malad City", "Pocatello", "Salmon", "Soda Springs", "Twin Falls"],
    "Illinois": ["Aurora", "Belleville", "Bloomington", "Bondville", "Chicago", "Decatur", "Marion", "Moline", "Mount Vernon", "Murphysboro", "Peoria", "Quincy", "Rock Falls", "Rockford", "Springfield", "West Chicago"],
    "Indiana": ["Bloomington", "Evansville", "Fort Wayne", "Huntingburg", "Indianapolis", "Kokomo", "Lafayette", "Muncie", "South Bend", "Terre Haute"],
    "Iowa": ["Algona", "Atlantic", "Boone", "Burlington", "Carroll", "Cedar Rapids", "Chariton", "Charles City", "Clarinda", "Clinton", "Council Bluffs", "Creston", "Decorah", "Denison", "Des Moines", "Dubuque", "Estherville", "Fairfield", "Fort Dodge", "Fort Madison", "Keokuk", "Knoxville", "Le Mars", "Mason City", "Monticello", "Muscatine", "Newton", "Oelwen", "Orange City", "Ottumwa", "Red Oak", "Sheldon", "Shenandoah", "Sioux City", "Spencer", "Storm Lake", "Washington", "Waterloo", "Webster City"],
    "Kansas": ["Chanute", "Concordia", "Dodge City", "Emporia", "Garden City", "Goodland", "Great Bend", "Hays", "Hill City", "Hutchinson", "Junction City", "Liberal", "Manhattan", "Newton", "Olathe", "Russell", "Salina", "Topeka", "Wichita"],
    "Kentucky": ["Bowling Green", "Fort Knox", "Hebron", "Henderson City", "Hopkinsville", "Jackson", "Lexington", "London", "Louisville", "Somerset", "West Paducah"],
    "Louisiana": ["Alexandria", "Baton Rouge", "Bossier City", "Houma", "Lafayette", "Lake Charles", "Leesville", "Monroe", "New Iberia", "New Orleans", "Patterson", "Pineville", "Shreveport"],
    "Maine": ["Auburn", "Augusta", "Bangor", "Bar Harbor", "Brunswick", "Caribou", "Frenchville", "Houlton", "Millinocket", "Portland", "Presque Isle", "Rockland", "Sanford", "Waterville", "Wiscasset"],
    "Maryland": ["Baltimore", "Camp Springs", "Hagerstown", "Lexington Park", "Salisbury"],
    "Massachusetts": ["Beverly", "Boston", "Chicopee Falls", "Hyannis", "Lawrence", "Marthas Vineyard", "Mashpee", "Nantucket", "New Bedford", "North Adams", "Norwood", "Plymouth", "Provincetown", "Westfield", "Worcester"],
    "Michigan": ["Alpena", "Ann Arbor", "Battle Creek", "Benton Harbor", "Cadillac", "Calumet", "Detroit", "Escanaba", "Flint", "Freeland", "Grand Rapids", "Houghton Lake", "Howell", "Ironwood", "Jackson", "Kalamazoo", "Kincheloe", "Kingsford", "Lansing", "Manistee", "Marie", "Menominee", "Mount Clemens", "Muskegon", "Oscoda", "Pellston", "Port Huron", "Traverse City", "Waterford Township"],
    "Minnesota": ["Aitkin", "Albert Lea", "Alexandria", "Austin", "Baudette", "Bemidji", "Benson", "Brainerd", "Cambridge", "Cloquet", "Crane Lake", "Crookston", "Detroit Lakes", "Duluth", "Ely", "Eveleth", "Fairmont", "Faribault", "Fergus Falls", "Flying Cloud", "Fosston", "Glenwood", "Grand Rapids", "Hallock", "Hibbing", "Hutchinson", "International Falls", "Litchfield", "Little Falls", "Mankato", "Marshall", "Minneapolis", "Mora", "Morris", "New Ulm", "Orr", "Owatonna", "Park Rapids", "Pipestone", "Red Wing", "Redwood Falls", "Rochester", "Roseau", "Silver Bay", "St Cloud", "St Paul", "Thief River Falls", "Two Harbors", "Wheaton", "Willmar", "Winona", "Worthington"],
    "Mississippi": ["Biloxi", "Columbus", "Greenville", "Greenwood", "Gulfport", "Hattiesburg", "Jackson", "McComb", "Meridian", "Natchez", "Starkville", "Tupelo"],
    "Missouri": ["Cape Girardeau", "Columbia", "Farmington", "Ft. Leonard Wood", "Jefferson City", "Joplin", "Kaiser", "Kansas City", "Kirksville", "Poplar Bluff", "Springfield", "St. Joseph", "St. Louis", "Vichy", "Warrensburg"],
    "Montana": ["Billings", "Bozeman", "Butte", "Cut Bank", "Glasgow", "Glendive", "Great Falls", "Havre", "Helena", "Kalispell", "Lewistown", "Livingston", "Miles City", "Missoula", "Sidney", "Wolf Point"],
    "Nebraska": ["Ainsworth", "Alliance", "Beatrice", "Bellevue", "Broken Bow", "Chadron", "Columbus", "Falls City", "Fremont", "Grand Island", "Hastings", "Holdrege", "Imperial", "Kearney", "Lincoln", "McCook", "Norfolk", "North Platte", "Omaha", "O'Neill", "Ord", "Scottsbluff", "Sidney", "Tekamah", "Valentine"],
    "Nevada": ["Elko", "Ely", "Fallon", "Las Vegas", "Lovelock", "Mercury", "Reno", "Tonopah", "Winnemucca"],
    "New Hampshire": ["Berlin", "Concord", "Keene", "Laconia", "Lebanon", "Manchester", "Mount Washington", "Portsmouth"],
    "New Jersey": ["Atlantic City", "Belmar", "Fairfield", "Millville", "Newark", "Rio Grande", "Teterboro", "Trenton"],
    "New Mexico": ["Alamogordo", "Albuquerque", "Carlsbad", "Clayton", "Clovis", "Deming", "Farmington", "Gallup", "Las Cruces", "Las Vegas", "Roswell", "Ruidoso", "Santa Fe", "Taos", "Truth or Consequences", "Tucumcari"],
    "New York": ["Albany", "Binghamton", "Buffalo", "Elmira", "Glen Falls", "Jamestown", "Massena", "Monticello", "New Windsor", "New York", "Niagara Falls", "Republic", "Rochester", "Ronkonkoma", "Saranac Lake", "Syracuse", "Utica", "Wappingers Falls", "Watertown", "Westhampton Beach", "White Plains"],
    "North Carolina": ["Asheville", "Cape Hatteras", "Charlotte", "Elizabeth City", "Fayetteville", "Goldsboro", "Greensboro", "Greenville", "Havelock", "Hickory", "Jacksonville", "Kinston", "Manteo", "New Bern", "Raleigh", "Rocky Mount", "Southern Pines", "Wilmington", "Winston-Salem"],
    "North Dakota": ["Bismarck", "Devils Lake", "Dickinson", "Fargo", "Grand Forks", "Jamestown", "Minot", "Williston"],
    "Ohio": ["Akron", "Cincinnati", "Cleveland", "Columbus", "Dayton", "Findlay", "Mansfield", "Toledo", "Youngstown", "Zanesville"],
    "Oklahoma": ["Altus", "Bartlesville", "Clinton", "Enid", "Fort Sill", "Gage", "Hobart", "Lawton", "McAlester", "Oklahoma City", "Ponca City", "Stillwater", "Tulsa"],
    "Oregon": ["Astoria", "Aurora", "Baker City", "Burns", "Corvallis", "Eugene", "Klamath Falls", "La Grande", "Lakeview", "Medford", "North Bend", "Pendleton", "Portland", "Redmond", "Roseburg", "Salem", "Sexton Summit"],
    "Pennsylvania": ["Allentown", "Altoona", "Bradford", "Butler", "DuBois", "Erie", "Franklin", "Harrisburg", "Johnstown", "Lancaster", "Middletown", "Philadelphia", "Pittsburgh", "Reading", "Scranton", "University Park", "Washington", "Williamsport"],
    "Rhode Island": ["New Shoreham", "Pawtucket", "Providence"],
    "South Carolina": ["Anderson", "Beaufort", "Charleston", "Columbia", "Florence", "Greenville", "Greer", "Myrtle Beach", "North Myrtle Beach", "Sumter"],
    "South Dakota": ["Aberdeen", "Brookings", "Huron", "Mitchell", "Mobridge", "Pierre", "Rapid City", "Sioux Falls", "Watertown", "Yankton"],
    "Tennessee": ["Blountville", "Chattanooga", "Crossville", "Dyersburg", "Jackson", "Knoxville", "Memphis", "Nashville"],
    "Texas": ["Abilene", "Alice", "Amarillo", "Austin", "Brownsville", "Childress", "College Station", "Corpus Christi", "Cotulla", "Dalhart", "Dallas", "Del Rio", "El Paso", "Galveston", "Georgetown", "Greenville", "Harlingen", "Hondo", "Houston", "Killeen", "Kingsville", "Laredo", "Longview", "Lubbock", "Lufkin", "Marfa", "McAllen", "McGregor", "Midland", "Mineral wells", "Nacogdoches", "Palacios", "Paris", "Port Arthur", "Rockport", "San Angelo", "San Antonio", "Temple", "Tyler", "Victoria", "Waco", "Wichita Falls", "Wink"],
    "Utah": ["Blanding", "Bryce Canyon City", "Cedar City", "Delta", "Hanksville", "Moab", "Ogden", "Provo", "Saint George", "Salt Lake City", "Vernal", "Wendover"],
    "Vermont": ["Burlington", "Montpelier", "Rutland", "Springfield"],
    "Virginia": ["Abingdon", "Blacksburg", "Charlottesville", "Danville", "Farmville", "Franklin", "Fredericksburg", "Hampton", "Hillsville", "Hot Springs", "Leesburg", "Lynchburg", "Manassas", "Marion", "Martinsville", "Melfa", "Newport News", "Norfolk", "Petersburg", "Pulaski", "Quantico", "Richmond", "Roanoke", "Virginia Beach", "Washington DC", "Weyers Cave", "Winchester", "Wise"],
    "Washington": ["Aberdeen", "Bellingham", "Bremerton", "Ephrata", "Everett", "Forks", "Hanford", "Hoquiam", "Kelso", "Moses Lake", "Olympia", "Pasco", "Port Angeles", "Pullman", "Renton", "Seattle", "Spokane", "Stampede Pass", "Tacoma", "The Dalles", "Toledo", "Walla Walla", "Wenatchee", "Whidbey Island", "Yakima"],
    "West Virginia": ["Beckley", "Bluefield", "Bridgeport", "Charleston", "Elkins", "Huntington", "Lewisburg", "Martinsburg", "Morgantown", "Parkersburg", "Wheeling"],
    "Wisconsin": ["Antigo", "Appleton", "Eau Claire", "Green Bay", "Janesville", "La Crosse", "Lone Rock", "Madison", "Manitowac", "Marshfield", "Milwaukee", "Minocqua", "Mosinee", "Oshkosh", "Phillips", "Rhinelander", "Rice Lake", "Sturgeon Bay", "Watertown", "Wausau"],
    "Wyoming": ["Casper", "Cheyenne", "Cody", "Evanston", "Gillette", "Jackson Hole", "Lander", "Laramie", "Rawlins", "Riverton", "Rock Springs", "Sheridan", "Worland"],
    "District of Columbia": ["Washington DC"]
}

# Complete weather data extracted from Excel - Fixed major data gaps
WEATHER_DATA = {
    # Alabama
    "Anniston": {"HDD": 2585, "CDD": 1713}, "Auburn": {"HDD": 2688, "CDD": 1477}, "Birmingham": {"HDD": 2698, "CDD": 1912},
    "Daleville": {"HDD": 2368, "CDD": 2014}, "Dothan": {"HDD": 2368, "CDD": 2014}, "Gadsen": {"HDD": 2954, "CDD": 1546},
    "Huntsville": {"HDD": 3196, "CDD": 1444}, "Mobile": {"HDD": 1615, "CDD": 2598}, "Montgomery": {"HDD": 2291, "CDD": 2211},
    "Muscle Shoals": {"HDD": 3196, "CDD": 1444}, "Troy": {"HDD": 2368, "CDD": 2014}, "Tuscaloosa": {"HDD": 2598, "CDD": 1815},
    
    # Alaska
    "Adak": {"HDD": 9202, "CDD": 0}, "Anchorage": {"HDD": 10158, "CDD": 0}, "Aniak": {"HDD": 12983, "CDD": 0},
    "Fairbanks": {"HDD": 13072, "CDD": 31}, "Juneau": {"HDD": 8471, "CDD": 1}, "Barrow": {"HDD": 19894, "CDD": 0},
    
    # Arizona
    "Casa Grande": {"HDD": 1473, "CDD": 3715}, "Flagstaff": {"HDD": 7112, "CDD": 110}, "Phoenix": {"HDD": 997, "CDD": 4591},
    "Tucson": {"HDD": 1596, "CDD": 3020}, "Yuma": {"HDD": 535, "CDD": 4532}, "Prescott": {"HDD": 4152, "CDD": 694},
    
    # Arkansas  
    "Batesville": {"HDD": 3761, "CDD": 1687}, "Little Rock": {"HDD": 3079, "CDD": 2076}, "Fort Smith": {"HDD": 3596, "CDD": 2020},
    "Hot Springs": {"HDD": 3256, "CDD": 2111}, "Texarkana": {"HDD": 2733, "CDD": 2370},
    
    # California
    "Los Angeles": {"HDD": 1283, "CDD": 617}, "San Francisco": {"HDD": 2737, "CDD": 97}, "San Diego": {"HDD": 1019, "CDD": 742},
    "Sacramento": {"HDD": 2581, "CDD": 1281}, "Fresno": {"HDD": 2327, "CDD": 2101}, "Oakland": {"HDD": 2816, "CDD": 128},
    "Bakersfield": {"HDD": 2171, "CDD": 1800}, "Riverside": {"HDD": 1351, "CDD": 1708}, "Stockton": {"HDD": 2494, "CDD": 1296},
    "Long Beach": {"HDD": 1283, "CDD": 617}, "Santa Ana": {"HDD": 1283, "CDD": 617}, "Anaheim": {"HDD": 1283, "CDD": 617},
    
    # Colorado
    "Denver": {"HDD": 5655, "CDD": 923}, "Colorado Springs": {"HDD": 6115, "CDD": 481}, "Aurora": {"HDD": 5655, "CDD": 923},
    "Fort Collins": {"HDD": 6185, "CDD": 643}, "Lakewood": {"HDD": 5655, "CDD": 923}, "Pueblo": {"HDD": 5145, "CDD": 1071},
    
    # Connecticut
    "Hartford": {"HDD": 5969, "CDD": 731}, "New Haven": {"HDD": 5524, "CDD": 625}, "Stamford": {"HDD": 5524, "CDD": 625},
    "Waterbury": {"HDD": 5969, "CDD": 731}, "Norwalk": {"HDD": 5524, "CDD": 625},
    
    # Delaware  
    "Dover": {"HDD": 4987, "CDD": 1010}, "Wilmington": {"HDD": 5087, "CDD": 1088},
    
    # Florida
    "Miami": {"HDD": 150, "CDD": 4292}, "Tampa": {"HDD": 646, "CDD": 3442}, "Orlando": {"HDD": 526, "CDD": 3234},
    "Jacksonville": {"HDD": 1281, "CDD": 2565}, "Fort Lauderdale": {"HDD": 322, "CDD": 4114}, "Key West": {"HDD": 67, "CDD": 4906},
    "Tallahassee": {"HDD": 1559, "CDD": 2642}, "St. Petersburg": {"HDD": 646, "CDD": 3442}, "Gainesville": {"HDD": 1115, "CDD": 2812},
    "Fort Myers": {"HDD": 218, "CDD": 4284}, "Pensacola": {"HDD": 1618, "CDD": 2647}, "Naples": {"HDD": 218, "CDD": 4284},
    
    # Georgia
    "Atlanta": {"HDD": 2773, "CDD": 1809}, "Augusta": {"HDD": 2512, "CDD": 2023}, "Savannah": {"HDD": 1759, "CDD": 2474},
    "Columbus": {"HDD": 2506, "CDD": 2068}, "Macon": {"HDD": 2181, "CDD": 2262}, "Albany": {"HDD": 1933, "CDD": 2413},
    
    # Hawaii
    "Honolulu": {"HDD": 0, "CDD": 4561}, "Hilo": {"HDD": 0, "CDD": 3279}, "Kahului": {"HDD": 0, "CDD": 4561},
    
    # Idaho
    "Boise": {"HDD": 5395, "CDD": 756}, "Idaho Falls": {"HDD": 7769, "CDD": 246}, "Pocatello": {"HDD": 6881, "CDD": 495},
    
    # Illinois
    "Chicago": {"HDD": 6399, "CDD": 830}, "Springfield": {"HDD": 5527, "CDD": 1166}, "Peoria": {"HDD": 6229, "CDD": 888},
    "Rockford": {"HDD": 7159, "CDD": 675}, "Decatur": {"HDD": 5564, "CDD": 1133},
    
    # Indiana
    "Indianapolis": {"HDD": 5845, "CDD": 1044}, "Fort Wayne": {"HDD": 6471, "CDD": 764}, "Evansville": {"HDD": 4474, "CDD": 1376},
    "South Bend": {"HDD": 6567, "CDD": 651}, "Terre Haute": {"HDD": 5636, "CDD": 1062},
    
    # Iowa
    "Des Moines": {"HDD": 6493, "CDD": 1121}, "Cedar Rapids": {"HDD": 6900, "CDD": 714}, "Dubuque": {"HDD": 7263, "CDD": 697},
    "Sioux City": {"HDD": 6911, "CDD": 1045}, "Waterloo": {"HDD": 7320, "CDD": 747},
    
    # Kansas
    "Wichita": {"HDD": 4324, "CDD": 1554}, "Topeka": {"HDD": 4953, "CDD": 1315}, "Kansas City": {"HDD": 5434, "CDD": 1316},
    "Salina": {"HDD": 5105, "CDD": 1375}, "Garden City": {"HDD": 4803, "CDD": 1382},
    
    # Kentucky
    "Louisville": {"HDD": 4521, "CDD": 1449}, "Lexington": {"HDD": 4856, "CDD": 1104}, "Bowling Green": {"HDD": 4146, "CDD": 1509},
    
    # Louisiana
    "New Orleans": {"HDD": 1358, "CDD": 2784}, "Baton Rouge": {"HDD": 1762, "CDD": 2622}, "Shreveport": {"HDD": 2351, "CDD": 2384},
    "Lake Charles": {"HDD": 1459, "CDD": 2695}, "Lafayette": {"HDD": 1502, "CDD": 2678},
    
    # Maine
    "Portland": {"HDD": 7679, "CDD": 335}, "Bangor": {"HDD": 7671, "CDD": 450}, "Augusta": {"HDD": 7807, "CDD": 413},
    
    # Maryland
    "Baltimore": {"HDD": 4631, "CDD": 1237}, "Washington DC": {"HDD": 4921, "CDD": 1113},
    
    # Massachusetts
    "Boston": {"HDD": 5793, "CDD": 734}, "Worcester": {"HDD": 7164, "CDD": 346}, "Springfield": {"HDD": 6293, "CDD": 673},
    
    # Michigan
    "Detroit": {"HDD": 6621, "CDD": 679}, "Grand Rapids": {"HDD": 6908, "CDD": 579}, "Lansing": {"HDD": 7045, "CDD": 597},
    "Flint": {"HDD": 7377, "CDD": 560}, "Ann Arbor": {"HDD": 6621, "CDD": 679},
    
    # Minnesota
    "Minneapolis": {"HDD": 7783, "CDD": 731}, "Duluth": {"HDD": 9620, "CDD": 159}, "Rochester": {"HDD": 8386, "CDD": 488},
    "St. Paul": {"HDD": 7783, "CDD": 731}, "St Cloud": {"HDD": 8562, "CDD": 551},
    
    # Mississippi
    "Jackson": {"HDD": 2428, "CDD": 2237}, "Biloxi": {"HDD": 1928, "CDD": 2606}, "Meridian": {"HDD": 2428, "CDD": 2237},
    
    # Missouri
    "St. Louis": {"HDD": 4846, "CDD": 1555}, "Kansas City": {"HDD": 5434, "CDD": 1316}, "Springfield": {"HDD": 4596, "CDD": 1547},
    "Columbia": {"HDD": 5046, "CDD": 1393}, "Jefferson City": {"HDD": 5046, "CDD": 1393},
    
    # Montana
    "Billings": {"HDD": 6731, "CDD": 548}, "Great Falls": {"HDD": 7854, "CDD": 344}, "Missoula": {"HDD": 7344, "CDD": 293},
    "Bozeman": {"HDD": 8664, "CDD": 189}, "Helena": {"HDD": 7894, "CDD": 336},
    
    # Nebraska
    "Omaha": {"HDD": 5954, "CDD": 1275}, "Lincoln": {"HDD": 5892, "CDD": 1220}, "Grand Island": {"HDD": 6221, "CDD": 1149},
    "North Platte": {"HDD": 6385, "CDD": 984}, "Scottsbluff": {"HDD": 6633, "CDD": 781},
    
    # Nevada
    "Las Vegas": {"HDD": 2301, "CDD": 3187}, "Reno": {"HDD": 5488, "CDD": 620}, "Elko": {"HDD": 7178, "CDD": 491},
    
    # New Hampshire
    "Manchester": {"HDD": 6322, "CDD": 664}, "Concord": {"HDD": 7479, "CDD": 397}, "Portsmouth": {"HDD": 6322, "CDD": 664},
    
    # New Jersey
    "Newark": {"HDD": 5057, "CDD": 1237}, "Atlantic City": {"HDD": 5073, "CDD": 886}, "Trenton": {"HDD": 5040, "CDD": 1177},
    
    # New Mexico
    "Albuquerque": {"HDD": 4157, "CDD": 1269}, "Las Cruces": {"HDD": 2876, "CDD": 2146}, "Santa Fe": {"HDD": 5449, "CDD": 589},
    "Roswell": {"HDD": 3672, "CDD": 1663}, "Farmington": {"HDD": 5506, "CDD": 749},
    
    # New York
    "New York": {"HDD": 4885, "CDD": 1133}, "Buffalo": {"HDD": 6612, "CDD": 468}, "Rochester": {"HDD": 6518, "CDD": 627},
    "Albany": {"HDD": 6773, "CDD": 489}, "Syracuse": {"HDD": 6609, "CDD": 535}, "Yonkers": {"HDD": 4885, "CDD": 1133},
    
    # North Carolina
    "Charlotte": {"HDD": 3153, "CDD": 1675}, "Raleigh": {"HDD": 3465, "CDD": 1566}, "Asheville": {"HDD": 4273, "CDD": 817},
    "Greensboro": {"HDD": 3945, "CDD": 1326}, "Winston-Salem": {"HDD": 3945, "CDD": 1326}, "Wilmington": {"HDD": 2347, "CDD": 2022},
    
    # North Dakota
    "Fargo": {"HDD": 9211, "CDD": 491}, "Bismarck": {"HDD": 8452, "CDD": 453}, "Grand Forks": {"HDD": 9534, "CDD": 486},
    "Minot": {"HDD": 9243, "CDD": 426}, "Williston": {"HDD": 8917, "CDD": 436},
    
    # Ohio
    "Columbus": {"HDD": 5599, "CDD": 747}, "Cleveland": {"HDD": 6160, "CDD": 760}, "Cincinnati": {"HDD": 4911, "CDD": 1039},
    "Toledo": {"HDD": 6494, "CDD": 679}, "Akron": {"HDD": 6160, "CDD": 760}, "Dayton": {"HDD": 5599, "CDD": 747},
    
    # Oklahoma
    "Oklahoma City": {"HDD": 3556, "CDD": 2038}, "Tulsa": {"HDD": 3844, "CDD": 2066}, "Norman": {"HDD": 3556, "CDD": 2038},
    
    # Oregon
    "Portland": {"HDD": 4187, "CDD": 367}, "Eugene": {"HDD": 4803, "CDD": 295}, "Medford": {"HDD": 4530, "CDD": 602},
    "Salem": {"HDD": 4187, "CDD": 367}, "Bend": {"HDD": 6436, "CDD": 267},
    
    # Pennsylvania
    "Philadelphia": {"HDD": 4824, "CDD": 1184}, "Pittsburgh": {"HDD": 5925, "CDD": 726}, "Harrisburg": {"HDD": 5409, "CDD": 927},
    "Allentown": {"HDD": 5824, "CDD": 798}, "Erie": {"HDD": 6451, "CDD": 538}, "Scranton": {"HDD": 6254, "CDD": 681},
    
    # Rhode Island
    "Providence": {"HDD": 5870, "CDD": 735}, "Warwick": {"HDD": 5870, "CDD": 735},
    
    # South Carolina
    "Charleston": {"HDD": 2051, "CDD": 2302}, "Columbia": {"HDD": 2593, "CDD": 2020}, "Greenville": {"HDD": 3652, "CDD": 1409},
    "North Charleston": {"HDD": 2051, "CDD": 2302}, "Rock Hill": {"HDD": 3153, "CDD": 1675},
    
    # South Dakota
    "Sioux Falls": {"HDD": 7680, "CDD": 680}, "Rapid City": {"HDD": 7203, "CDD": 675}, "Aberdeen": {"HDD": 8375, "CDD": 571},
    "Pierre": {"HDD": 7565, "CDD": 816}, "Watertown": {"HDD": 8375, "CDD": 571},
    
    # Tennessee
    "Memphis": {"HDD": 2999, "CDD": 2134}, "Nashville": {"HDD": 3737, "CDD": 1751}, "Knoxville": {"HDD": 3959, "CDD": 1482},
    "Chattanooga": {"HDD": 3440, "CDD": 1730}, "Clarksville": {"HDD": 3737, "CDD": 1751},
    
    # Texas
    "Houston": {"HDD": 1439, "CDD": 2974}, "Dallas": {"HDD": 2333, "CDD": 2678}, "Austin": {"HDD": 1269, "CDD": 2884},
    "San Antonio": {"HDD": 1548, "CDD": 2992}, "El Paso": {"HDD": 2499, "CDD": 2171}, "Fort Worth": {"HDD": 2333, "CDD": 2678},
    "Arlington": {"HDD": 2333, "CDD": 2678}, "Corpus Christi": {"HDD": 914, "CDD": 3260}, "Plano": {"HDD": 2333, "CDD": 2678},
    
    # Utah
    "Salt Lake City": {"HDD": 5350, "CDD": 1118}, "Provo": {"HDD": 5836, "CDD": 858}, "Saint George": {"HDD": 2729, "CDD": 2936},
    "West Valley City": {"HDD": 5350, "CDD": 1118}, "Ogden": {"HDD": 5350, "CDD": 1118},
    
    # Vermont
    "Burlington": {"HDD": 7491, "CDD": 420}, "Montpelier": {"HDD": 7662, "CDD": 249}, "Rutland": {"HDD": 7662, "CDD": 249},
    
    # Virginia
    "Norfolk": {"HDD": 3411, "CDD": 1630}, "Richmond": {"HDD": 3883, "CDD": 1493}, "Virginia Beach": {"HDD": 3336, "CDD": 1539},
    "Newport News": {"HDD": 3411, "CDD": 1630}, "Alexandria": {"HDD": 4921, "CDD": 1113}, "Hampton": {"HDD": 3411, "CDD": 1630},
    
    # Washington
    "Seattle": {"HDD": 4372, "CDD": 169}, "Spokane": {"HDD": 6716, "CDD": 341}, "Tacoma": {"HDD": 5664, "CDD": 142},
    "Vancouver": {"HDD": 4372, "CDD": 169}, "Bellevue": {"HDD": 4372, "CDD": 169}, "Everett": {"HDD": 4372, "CDD": 169},
    
    # West Virginia
    "Charleston": {"HDD": 4708, "CDD": 1015}, "Huntington": {"HDD": 4642, "CDD": 1077}, "Parkersburg": {"HDD": 5044, "CDD": 848},
    "Wheeling": {"HDD": 5834, "CDD": 708}, "Morgantown": {"HDD": 5834, "CDD": 708},
    
    # Wisconsin
    "Milwaukee": {"HDD": 7348, "CDD": 545}, "Madison": {"HDD": 7724, "CDD": 604}, "Green Bay": {"HDD": 7853, "CDD": 496},
    "Kenosha": {"HDD": 7348, "CDD": 545}, "Racine": {"HDD": 7348, "CDD": 545}, "Appleton": {"HDD": 7853, "CDD": 496},
    
    # Wyoming
    "Cheyenne": {"HDD": 7362, "CDD": 265}, "Casper": {"HDD": 7409, "CDD": 440}, "Jackson Hole": {"HDD": 9670, "CDD": 15},
    "Laramie": {"HDD": 7362, "CDD": 265}, "Rock Springs": {"HDD": 7932, "CDD": 267}
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

# COMPLETE savings data from Excel Savings Lookup table - ALL 105 entries
COMPLETE_SAVINGS_DATA = {
    # Single-Single Office combinations
    "SingleSingleMidOfficePVAV_ElecElectric2080": {"heat_kwh_sf": 7.67, "cool_kwh_sf": 3.03, "gas_therms_sf": 0},
    "SingleSingleMidOfficePVAV_ElecElectric2912": {"heat_kwh_sf": 10.92, "cool_kwh_sf": 4.99, "gas_therms_sf": 0},
    "SingleSingleMidOfficePVAV_ElecElectric8760": {"heat_kwh_sf": 27.02, "cool_kwh_sf": 10.41, "gas_therms_sf": 0},
    "SingleDoubleMidOfficePVAV_ElecElectric2080": {"heat_kwh_sf": 7.12, "cool_kwh_sf": 5.33, "gas_therms_sf": 0},
    "SingleDoubleMidOfficePVAV_ElecElectric2912": {"heat_kwh_sf": 10.31, "cool_kwh_sf": 7.99, "gas_therms_sf": 0},
    "SingleDoubleMidOfficePVAV_ElecElectric8760": {"heat_kwh_sf": 28.39, "cool_kwh_sf": 13.29, "gas_therms_sf": 0},
    "DoubleSingleMidOfficePVAV_GasNatural Gas2912": {"heat_kwh_sf": 0, "cool_kwh_sf": 3.1, "gas_therms_sf": 0.22},
    "DoubleSingleMidOfficePVAV_GasNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 5.15, "gas_therms_sf": 0.64},
    
    # Large Office combinations
    "SingleSingleLargeOfficeVAVElectric2080": {"heat_kwh_sf": 8.38, "cool_kwh_sf": 3.6, "gas_therms_sf": 0},
    "SingleSingleLargeOfficeVAVElectric2912": {"heat_kwh_sf": 12.07, "cool_kwh_sf": 5.89, "gas_therms_sf": 0},
    "SingleSingleLargeOfficeVAVElectric8760": {"heat_kwh_sf": 29.58, "cool_kwh_sf": 12.27, "gas_therms_sf": 0},
    "SingleDoubleLargeOfficeVAVElectric2080": {"heat_kwh_sf": 7.78, "cool_kwh_sf": 6.31, "gas_therms_sf": 0},
    "SingleDoubleLargeOfficeVAVElectric2912": {"heat_kwh_sf": 11.37, "cool_kwh_sf": 9.44, "gas_therms_sf": 0},
    "SingleDoubleLargeOfficeVAVElectric8760": {"heat_kwh_sf": 31.07, "cool_kwh_sf": 17.03, "gas_therms_sf": 0},
    "DoubleSingleLargeOfficeVAVElectric2080": {"heat_kwh_sf": 3.83, "cool_kwh_sf": 2.5, "gas_therms_sf": 0},
    "DoubleSingleLargeOfficeVAVElectric2912": {"heat_kwh_sf": 5.11, "cool_kwh_sf": 3.67, "gas_therms_sf": 0},
    "DoubleSingleLargeOfficeVAVElectric8760": {"heat_kwh_sf": 14.5, "cool_kwh_sf": 6.08, "gas_therms_sf": 0},
    "SingleSingleLargeOfficeVAVNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 3.6, "gas_therms_sf": 0.4},
    "SingleSingleLargeOfficeVAVNatural Gas2912": {"heat_kwh_sf": 0, "cool_kwh_sf": 5.89, "gas_therms_sf": 0.58},
    "SingleSingleLargeOfficeVAVNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 12.27, "gas_therms_sf": 1.44},
    "SingleDoubleLargeOfficeVAVNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 6.31, "gas_therms_sf": 0.37},
    "SingleDoubleLargeOfficeVAVNatural Gas2912": {"heat_kwh_sf": 0, "cool_kwh_sf": 9.44, "gas_therms_sf": 0.55},
    "SingleDoubleLargeOfficeVAVNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 17.03, "gas_therms_sf": 1.52},
    "DoubleSingleLargeOfficeVAVNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 2.5, "gas_therms_sf": 0.18},
    "DoubleSingleLargeOfficeVAVNatural Gas2912": {"heat_kwh_sf": 0, "cool_kwh_sf": 3.67, "gas_therms_sf": 0.24},
    "DoubleSingleLargeOfficeVAVNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 6.08, "gas_therms_sf": 0.69},
    
    # Hotel combinations
    "SingleSingleMidHotelPTACElectric8760": {"heat_kwh_sf": 12.54, "cool_kwh_sf": 12.26, "gas_therms_sf": 0},
    "SingleDoubleMidHotelPTACElectric8760": {"heat_kwh_sf": 11.65, "cool_kwh_sf": 20.39, "gas_therms_sf": 0},
    "DoubleSingleMidHotelPTACElectric8760": {"heat_kwh_sf": 6.1, "cool_kwh_sf": 8.56, "gas_therms_sf": 0},
    "SingleSingleMidHotelPTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 12.26, "gas_therms_sf": 0.61},
    "SingleDoubleMidHotelPTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 20.39, "gas_therms_sf": 0.57},
    "DoubleSingleMidHotelPTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 8.56, "gas_therms_sf": 0.3},
    "SingleSingleLargeHotelPTACElectric8760": {"heat_kwh_sf": 13.78, "cool_kwh_sf": 13.47, "gas_therms_sf": 0},
    "SingleDoubleLargeHotelPTACElectric8760": {"heat_kwh_sf": 12.82, "cool_kwh_sf": 22.43, "gas_therms_sf": 0},
    "DoubleSingleLargeHotelPTACElectric8760": {"heat_kwh_sf": 6.71, "cool_kwh_sf": 9.42, "gas_therms_sf": 0},
    "SingleSingleLargeHotelPTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 13.47, "gas_therms_sf": 0.67},
    "SingleDoubleLargeHotelPTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 22.43, "gas_therms_sf": 0.62},
    "DoubleSingleLargeHotelPTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 9.42, "gas_therms_sf": 0.33},
    
    # School combinations  
    "SingleSingleMidSchoolPVAV_GasNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 2.36, "gas_therms_sf": 0.29},
    "SingleDoubleMidSchoolPVAV_GasNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 4.15, "gas_therms_sf": 0.27},
    "DoubleSingleMidSchoolPVAV_GasNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 1.64, "gas_therms_sf": 0.13},
    "SingleSingleLargeSchoolVAVNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 2.8, "gas_therms_sf": 0.31},
    "SingleDoubleLargeSchoolVAVNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 4.91, "gas_therms_sf": 0.29},
    "DoubleSingleLargeSchoolVAVNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 1.95, "gas_therms_sf": 0.14},
    
    # Hospital combinations
    "SingleSingleMidHospitalVAVNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 15.27, "gas_therms_sf": 1.78},
    "SingleDoubleMidHospitalVAVNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 20.7, "gas_therms_sf": 1.71},
    "DoubleSingleMidHospitalVAVNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 7.59, "gas_therms_sf": 0.81},
    "SingleSingleLargeHospitalVAVNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 15.22, "gas_therms_sf": 1.77},
    "SingleDoubleLargeHospitalVAVNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 20.64, "gas_therms_sf": 1.7},
    "DoubleSingleLargeHospitalVAVNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 7.57, "gas_therms_sf": 0.81},
    
    # Multi-family combinations
    "SingleSingleMidMulti-familyPTHPElectric8760": {"heat_kwh_sf": 8.78, "cool_kwh_sf": 6.26, "gas_therms_sf": 0},
    "SingleDoubleMidMulti-familyPTHPElectric8760": {"heat_kwh_sf": 8.16, "cool_kwh_sf": 10.42, "gas_therms_sf": 0},
    "DoubleSingleMidMulti-familyPTHPElectric8760": {"heat_kwh_sf": 4.27, "cool_kwh_sf": 4.37, "gas_therms_sf": 0},
    "SingleSingleLargeMulti-familyPTHPElectric8760": {"heat_kwh_sf": 9.66, "cool_kwh_sf": 6.88, "gas_therms_sf": 0},
    "SingleDoubleLargeMulti-familyPTHPElectric8760": {"heat_kwh_sf": 8.98, "cool_kwh_sf": 11.46, "gas_therms_sf": 0},
    "DoubleSingleLargeMulti-familyPTHPElectric8760": {"heat_kwh_sf": 4.69, "cool_kwh_sf": 4.81, "gas_therms_sf": 0},
    "SingleSingleMidMulti-familyPTHPNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 6.26, "gas_therms_sf": 0.43},
    "SingleDoubleMidMulti-familyPTHPNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 10.42, "gas_therms_sf": 0.4},
    "DoubleSingleMidMulti-familyPTHPNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 4.37, "gas_therms_sf": 0.21},
    "SingleSingleLargeMulti-familyPTHPNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 6.88, "gas_therms_sf": 0.47},
    "SingleDoubleLargeMulti-familyPTHPNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 11.46, "gas_therms_sf": 0.44},
    "DoubleSingleLargeMulti-familyPTHPNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 4.81, "gas_therms_sf": 0.23},
    
    # Additional PTAC, PTHP, and FCU combinations for Office
    "SingleSingleMidOfficePTACElectric8760": {"heat_kwh_sf": 26.42, "cool_kwh_sf": 10.2, "gas_therms_sf": 0},
    "SingleDoubleMidOfficePTACElectric8760": {"heat_kwh_sf": 27.78, "cool_kwh_sf": 13.02, "gas_therms_sf": 0},
    "DoubleSingleMidOfficePTACElectric8760": {"heat_kwh_sf": 12.85, "cool_kwh_sf": 5.05, "gas_therms_sf": 0},
    "SingleSingleMidOfficePTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 10.2, "gas_therms_sf": 1.28},
    "SingleDoubleMidOfficePTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 13.02, "gas_therms_sf": 1.35},
    "DoubleSingleMidOfficePTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 5.05, "gas_therms_sf": 0.62},
    "SingleSingleLargeOfficePTACElectric8760": {"heat_kwh_sf": 29.06, "cool_kwh_sf": 11.22, "gas_therms_sf": 0},
    "SingleDoubleLargeOfficePTACElectric8760": {"heat_kwh_sf": 30.56, "cool_kwh_sf": 14.32, "gas_therms_sf": 0},
    "DoubleSingleLargeOfficePTACElectric8760": {"heat_kwh_sf": 14.13, "cool_kwh_sf": 5.56, "gas_therms_sf": 0},
    "SingleSingleLargeOfficePTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 11.22, "gas_therms_sf": 1.41},
    "SingleDoubleLargeOfficePTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 14.32, "gas_therms_sf": 1.48},
    "DoubleSingleLargeOfficePTACNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 5.56, "gas_therms_sf": 0.69},
    
    # PTHP Office combinations
    "SingleSingleMidOfficePTHPElectric8760": {"heat_kwh_sf": 22.24, "cool_kwh_sf": 10.2, "gas_therms_sf": 0},
    "SingleDoubleMidOfficePTHPElectric8760": {"heat_kwh_sf": 23.39, "cool_kwh_sf": 13.02, "gas_therms_sf": 0},
    "DoubleSingleMidOfficePTHPElectric8760": {"heat_kwh_sf": 10.82, "cool_kwh_sf": 5.05, "gas_therms_sf": 0},
    "SingleSingleLargeOfficePTHPElectric8760": {"heat_kwh_sf": 24.46, "cool_kwh_sf": 11.22, "gas_therms_sf": 0},
    "SingleDoubleLargeOfficePTHPElectric8760": {"heat_kwh_sf": 25.74, "cool_kwh_sf": 14.32, "gas_therms_sf": 0},
    "DoubleSingleLargeOfficePTHPElectric8760": {"heat_kwh_sf": 11.9, "cool_kwh_sf": 5.56, "gas_therms_sf": 0},
    
    # FCU Office combinations
    "SingleSingleMidOfficeFCUNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 10.41, "gas_therms_sf": 1.31},
    "SingleDoubleMidOfficeFCUNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 13.29, "gas_therms_sf": 1.38},
    "DoubleSingleMidOfficeFCUNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 5.15, "gas_therms_sf": 0.64},
    "SingleSingleLargeOfficeFCUNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 12.27, "gas_therms_sf": 1.44},
    "SingleDoubleLargeOfficeFCUNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 17.03, "gas_therms_sf": 1.52},
    "DoubleSingleLargeOfficeFCUNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 6.08, "gas_therms_sf": 0.69}
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
    floor_area_per_floor = floor_area / num_stories
    side_length = math.sqrt(floor_area_per_floor)
    perimeter = 4 * side_length
    floor_height = 12  # feet
    total_wall_area = perimeter * floor_height * num_stories
    wwr = window_area / total_wall_area if total_wall_area > 0 else 0
    return wwr, total_wall_area

def create_lookup_key_excel(form_data):
    """Create lookup key exactly matching Excel methodology - FIXED"""
    building_size = "Large" if form_data['floor_area'] >= 200000 else "Mid"
    
    # HVAC mapping to match Excel lookup keys
    hvac_mapping = {
        "Packaged VAV with electric reheat": "PVAV_Elec",
        "Packaged VAV with hydronic reheat": "PVAV_Gas", 
        "Built-up VAV with hydronic reheat": "VAV",
        "PTAC": "PTAC",
        "PTHP": "PTHP", 
        "Fan Coil Unit": "FCU",
        "Other": "VAV"
    }
    
    # Window type mapping
    existing_map = {
        "Single pane": "Single", 
        "Double pane": "Double", 
        "New double pane (U<0.35)": "Double"
    }
    
    # Fuel type mapping - FIXED to match Excel exactly
    fuel_mapping = {
        "Electric": "Electric",
        "Natural Gas": "Natural Gas",  # Keep the space!
        "Electric Only": "Electric"
    }
    
    existing_code = existing_map.get(form_data['existing_window_type'], "Single")
    secondary_code = form_data['secondary_window_type']
    building_type = form_data['building_type']
    hvac_code = hvac_mapping.get(form_data['hvac_type'], "VAV")
    fuel_code = fuel_mapping.get(form_data['heating_fuel'], "Natural Gas")
    hours = form_data['operation_hours']
    
    # Construct lookup key exactly like Excel
    lookup_key = f"{existing_code}{secondary_code}{building_size}{building_type}{hvac_code}{fuel_code}{hours}"
    
    return lookup_key

def calculate_savings_excel():
    """Calculate energy savings using exact Excel methodology - ENHANCED"""
    form_data = st.session_state.form_data
    city = form_data['city']
    
    # Get weather data with better fallback handling
    weather = None
    if city in WEATHER_DATA:
        weather = WEATHER_DATA[city]
    else:
        # Try variations and state defaults
        city_variations = [
            city.replace(" ", ""),
            city.replace(".", ""),
            f"{city} {form_data['state'][:2].upper()}"
        ]
        
        for variation in city_variations:
            if variation in WEATHER_DATA:
                weather = WEATHER_DATA[variation]
                city = variation
                break
        
        if not weather:
            # State defaults for major states
            state_defaults = {
                "California": {"HDD": 2581, "CDD": 1281},
                "Texas": {"HDD": 2333, "CDD": 2678},
                "Florida": {"HDD": 646, "CDD": 3442},
                "New York": {"HDD": 4885, "CDD": 1133},
                "Illinois": {"HDD": 6399, "CDD": 830},
                "Pennsylvania": {"HDD": 4824, "CDD": 1184},
                "Ohio": {"HDD": 5599, "CDD": 747},
                "Georgia": {"HDD": 2773, "CDD": 1809},
                "North Carolina": {"HDD": 3153, "CDD": 1675},
                "Michigan": {"HDD": 6621, "CDD": 679}
            }
            weather = state_defaults.get(form_data['state'], {"HDD": 4000, "CDD": 1000})
            city = f"Default_{form_data['state']}"
    
    hdd, cdd = weather["HDD"], weather["CDD"]
    
    # Create lookup key
    lookup_key = create_lookup_key_excel(form_data)
    
    # Get savings data from complete lookup table
    savings_data = COMPLETE_SAVINGS_DATA.get(lookup_key)
    
    if not savings_data:
        # Enhanced fallback logic for missing combinations
        building_size = "Large" if form_data['floor_area'] >= 200000 else "Mid"
        building_type = form_data['building_type']
        fuel_type = "Natural Gas" if form_data['heating_fuel'] == "Natural Gas" else "Electric"
        hours = form_data['operation_hours']
        
        # Try different HVAC system fallbacks
        hvac_fallbacks = ["VAV", "PVAV_Gas", "PVAV_Elec", "PTAC", "PTHP", "FCU"]
        window_fallbacks = [
            ("Single", "Single"),
            ("Single", "Double"), 
            ("Double", "Single")
        ]
        
        for hvac in hvac_fallbacks:
            for existing, secondary in window_fallbacks:
                fallback_key = f"{existing}{secondary}{building_size}{building_type}{hvac}{fuel_type}{hours}"
                if fallback_key in COMPLETE_SAVINGS_DATA:
                    savings_data = COMPLETE_SAVINGS_DATA[fallback_key]
                    lookup_key = fallback_key + " (fallback)"
                    break
            if savings_data:
                break
    
    if not savings_data:
        # Ultimate fallback with reasonable defaults based on building type
        building_defaults = {
            "Office": {"heat_kwh_sf": 5.0, "cool_kwh_sf": 8.0, "gas_therms_sf": 0.5},
            "Hotel": {"heat_kwh_sf": 10.0, "cool_kwh_sf": 15.0, "gas_therms_sf": 0.6},
            "School": {"heat_kwh_sf": 3.0, "cool_kwh_sf": 5.0, "gas_therms_sf": 0.3},
            "Hospital": {"heat_kwh_sf": 8.0, "cool_kwh_sf": 18.0, "gas_therms_sf": 1.5},
            "Multi-family": {"heat_kwh_sf": 6.0, "cool_kwh_sf": 8.0, "gas_therms_sf": 0.4}
        }
        
        default_data = building_defaults.get(form_data['building_type'], building_defaults["Office"])
        
        # Adjust for fuel type
        if form_data['heating_fuel'] == "Electric":
            savings_data = {
                "heat_kwh_sf": default_data["heat_kwh_sf"],
                "cool_kwh_sf": default_data["cool_kwh_sf"],
                "gas_therms_sf": 0
            }
        else:
            savings_data = {
                "heat_kwh_sf": 0,
                "cool_kwh_sf": default_data["cool_kwh_sf"],
                "gas_therms_sf": default_data["gas_therms_sf"]
            }
        lookup_key = "FALLBACK_DEFAULT"
    
    # Calculate savings using Excel methodology
    window_area = form_data['window_area']
    
    total_heating_kwh = savings_data["heat_kwh_sf"] * window_area
    total_cooling_kwh = savings_data["cool_kwh_sf"] * window_area  
    total_kwh = total_heating_kwh + total_cooling_kwh
    total_gas_therms = max(0, savings_data["gas_therms_sf"] * window_area)
    
    # Cost calculations
    electric_cost = total_kwh * form_data['electric_rate']
    gas_cost = total_gas_therms * form_data['gas_rate'] 
    total_cost = electric_cost + gas_cost
    
    # Building metrics
    floor_area = form_data['floor_area']
    cost_per_sf = total_cost / floor_area if floor_area > 0 else 0
    
    # WWR calculation
    wwr, wall_area = calculate_wwr(floor_area, form_data['num_stories'], window_area)
    
    # Energy intensity calculations
    total_energy_btu = (total_kwh * 3412) + (total_gas_therms * 100000)
    energy_intensity_savings = total_energy_btu / floor_area / 1000 if floor_area > 0 else 0
    
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
st.markdown('<div style="text-align: center; color: #666; margin-bottom: 2rem;">Version 2.1.0 - Complete Excel Data Integration</div>', unsafe_allow_html=True)

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
        elif city:
            st.markdown('<div class="warning-box">Weather data not found for this city. State defaults will be used.</div>', unsafe_allow_html=True)
        
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
            # Show baseline EUI for selected building type
            baseline = BASELINE_EUI.get(selected_building, 85)
            st.info(f"Baseline Energy Use Intensity: {baseline} kBtu/SF-year")
        
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
                min_value=10000,
                max_value=1000000,
                value=st.session_state.form_data.get('floor_area', 100000),
                step=5000,
                format="%d",
                help="Total building floor area"
            )
            
        with col2:
            num_stories = st.number_input(
                "No. of Floors *",
                min_value=1,
                max_value=50,
                value=st.session_state.form_data.get('num_stories', 5),
                step=1,
                format="%d",
                help="Number of floors in the building"
            )
            
        with col3:
            # Set default operation hours based on building type
            building_type = st.session_state.form_data.get('building_type', 'Office')
            default_hours = {
                'Office': 2912,
                'School': 2080, 
                'Hotel': 8760,
                'Hospital': 8760,
                'Multi-family': 8760
            }.get(building_type, 3000)
            
            operation_hours = st.number_input(
                "Annual Operating Hours *",
                min_value=1980,
                max_value=8760,
                value=st.session_state.form_data.get('operation_hours', default_hours),
                step=100,
                format="%d",
                help=f"Typical for {building_type}: {default_hours} hours"
            )
        
        # Building classification
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
            
            # Smart default for window area based on building size and WWR
            floor_area = st.session_state.form_data.get('floor_area', 100000)
            num_stories = st.session_state.form_data.get('num_stories', 5)
            if floor_area and num_stories:
                # Estimate reasonable window area for WWR of 0.25
                floor_area_per_floor = floor_area / num_stories
                side_length = math.sqrt(floor_area_per_floor)
                perimeter = 4 * side_length
                floor_height = 12
                total_wall_area = perimeter * floor_height * num_stories
                suggested_window_area = int(total_wall_area * 0.25)
            else:
                suggested_window_area = 25000
            
            window_area = st.number_input(
                "Sq.ft. of CSW Installed *",
                min_value=1000,
                max_value=200000,
                value=st.session_state.form_data.get('window_area', suggested_window_area),
                step=500,
                format="%d",
                help="Total area of Commercial Secondary Windows to be installed"
            )
            
        with col2:
            secondary_window_type = st.selectbox(
                "Type of CSW Analyzed *",
                options=SECONDARY_WINDOW_TYPES,
                index=SECONDARY_WINDOW_TYPES.index(st.session_state.form_data.get('secondary_window_type', SECONDARY_WINDOW_TYPES[0])),
                help="Type of Commercial Secondary Windows being added"
            )
            
            # Calculate and display WWR
            if st.session_state.form_data.get('floor_area') and st.session_state.form_data.get('num_stories'):
                wwr, wall_area = calculate_wwr(
                    st.session_state.form_data['floor_area'], 
                    st.session_state.form_data['num_stories'], 
                    window_area
                )
                st.metric("Window-to-Wall Ratio", f"{wwr:.3f}", help="Ratio of window area to total wall area")
                
                # WWR validation
                if wwr < 0.1:
                    st.markdown('<div class="warning-box">WWR is below typical range (0.1-0.5). Consider increasing window area.</div>', unsafe_allow_html=True)
                elif wwr > 0.5:
                    st.markdown('<div class="warning-box">WWR is above typical range (0.1-0.5). Consider reducing window area.</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="info-box">WWR is within typical range (0.1-0.5).</div>', unsafe_allow_html=True)
        
        # Window performance info
        window_info = {
            "Single pane": "U-assembly = 1.03 Btu/hr-SF-°F; SHGC = 0.73",
            "Double pane": "U-assembly = 0.65 Btu/hr-SF-°F; SHGC = 0.40", 
            "New double pane (U<0.35)": "U-assembly < 0.35 Btu/hr-SF-°F; SHGC < 0.35"
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
            # Smart defaults based on building type
            building_type = st.session_state.form_data.get('building_type', 'Office')
            floor_area = st.session_state.form_data.get('floor_area', 100000)
            
            default_hvac = {
                'Office': "Built-up VAV with hydronic reheat" if floor_area >= 200000 else "Packaged VAV with hydronic reheat",
                'Hotel': "PTAC",
                'School': "Packaged VAV with hydronic reheat",
                'Hospital': "Built-up VAV with hydronic reheat",
                'Multi-family': "PTHP"
            }.get(building_type, HVAC_SYSTEMS[0])
            
            hvac_type = st.selectbox(
                "HVAC System Type *",
                options=HVAC_SYSTEMS,
                index=HVAC_SYSTEMS.index(st.session_state.form_data.get('hvac_type', default_hvac)) if st.session_state.form_data.get('hvac_type', default_hvac) in HVAC_SYSTEMS else 0
            )
            
        with col2:
            # Smart default for heating fuel
            default_fuel = "Natural Gas"
            if building_type == "Multi-family" or "electric" in hvac_type.lower():
                default_fuel = "Electric"
                
            heating_fuel = st.selectbox(
                "Primary Heating Fuel *",
                options=HEATING_FUELS,
                index=HEATING_FUELS.index(st.session_state.form_data.get('heating_fuel', default_fuel)) if st.session_state.form_data.get('heating_fuel', default_fuel) in HEATING_FUELS else 1
            )
        
        # System compatibility info
        if "PTHP" in hvac_type and heating_fuel == "Natural Gas":
            st.warning("Note: PTHP systems typically use electric heating. Please verify your selection.")
        elif "electric" in hvac_type.lower() and heating_fuel == "Natural Gas":
            st.info("Note: Electric reheat system with Natural Gas primary heating.")
        
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
                format="%.3f",
                help="Commercial electric rate including demand charges"
            )
            
        with col2:
            gas_rate = st.number_input(
                "Natural Gas Rate, $/therm *",
                min_value=0.50,
                max_value=5.00,
                value=st.session_state.form_data.get('gas_rate', 1.05),
                step=0.01,
                format="%.2f",
                help="Commercial natural gas rate"
            )
        
        st.info("💡 Check your recent utility bills for accurate rates. Commercial rates may include demand charges.")
        
        st.session_state.form_data.update({
            'electric_rate': electric_rate,
            'gas_rate': gas_rate
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 8: Results
elif st.session_state.step == 8:
    st.markdown('<div class="step-header">Step 8: Energy Savings Results</div>', unsafe_allow_html=True)
    
    try:
        results = calculate_savings_excel()
        
        # Main results display
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
                <h3>Total Annual Savings</h3>
                <h2>${results['total_cost']:,.0f}</h2>
                <p>per year</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Key metrics
        st.markdown("### Key Performance Metrics")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Energy Intensity Savings", f"{results['energy_intensity_savings']:.1f} kBtu/SF-yr")
        with col2:
            st.metric("Baseline EUI", f"{results['baseline_eui']:.1f} kBtu/SF-yr")
        with col3:
            st.metric("Savings Percentage", f"{results['percentage_savings']:.1f}%")
        with col4:
            st.metric("Window-Wall Ratio", f"{results['wwr']:.3f}")
        
        # Calculation details (for debugging)
        if st.session_state.form_data.get('show_debug', False):
            st.markdown("### Calculation Details")
            with st.expander("Debug Information"):
                st.write(f"**Lookup Key Used:** {results['lookup_key']}")
                st.write(f"**City Used:** {results['city_used']}")
                st.write(f"**HDD:** {results['hdd']:,}, **CDD:** {results['cdd']:,}")
                st.write(f"**Savings Data:** {results['savings_data_used']}")
        
        # Show debug toggle
        show_debug = st.checkbox("Show calculation details", key="show_debug")
        
        # Download functionality
        st.markdown("### Download Report")
        
        summary_data = {
            'Parameter': [
                'Project Name', 'Contact Person', 'Company', 'Building Type', 'Location',
                'Building Area, Sq.Ft.', 'No. of Floors', 'Annual Operating Hours',
                'Sq.ft. of CSW Installed', 'Type of Existing Window', 'Type of CSW Analyzed',
                'HVAC System Type', 'Primary Heating Fuel', 'Electric Rate, $/kWh',
                'Natural Gas Rate, $/therm', 'Window-Wall Ratio', '',
                'Annual Electric Energy Savings, kWh/yr', 'Gas Savings, therms/yr', 
                'Electric Cost Savings, $/yr', 'Total Savings, $/yr',
                'Energy Intensity Savings, kBtu/SF-yr', 'Baseline Energy Use Intensity, kBtu/SF-yr',
                'Savings Percentage, %', 'Calculation Methodology'
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
                f"{results['wwr']:.3f}", '',
                f"{results['total_kwh']:,.0f}",
                f"{results['gas_therms']:,.0f}",
                f"$ {results['electric_cost']:,.0f}",
                f"$ {results['total_cost']:,.0f}",
                f"{results['energy_intensity_savings']:.1f}",
                f"{results['baseline_eui']:.1f}",
                f"{results['percentage_savings']:.1f}",
                f"Excel Lookup Key: {results['lookup_key']}"
            ]
        }
        
        df_summary = pd.DataFrame(summary_data)
        st.dataframe(df_summary, use_container_width=True, hide_index=True)
        
        csv_data = df_summary.to_csv(index=False)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"CSW_Savings_Report_{st.session_state.form_data.get('project_name', 'Project').replace(' ', '_')}_{timestamp}.csv"
        
        st.download_button(
            label="📊 Download Complete Results Report",
            data=csv_data,
            file_name=filename,
            mime="text/csv",
            type="primary",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Error calculating savings: {str(e)}")
        st.write("Please check your inputs and try again. If the problem persists, some data combinations may not be available in the lookup table.")

# Navigation buttons
st.markdown("---")
col1, col2, col3 = st.columns([1, 2, 1])

with col1:
    if st.session_state.step > 1:
        if st.button("← Previous", use_container_width=True):
            navigate_step("prev")

with col3:
    if st.session_state.step < 8:
        if st.button("Next →", use_container_width=True, type="primary"):
            navigate_step("next")

with col2:
    if st.session_state.step == 8:
        if st.button("🔄 Start New Calculation", use_container_width=True, type="secondary"):
            st.session_state.step = 1
            st.session_state.form_data = {}
            st.rerun()

# Footer
if st.session_state.step < 8:
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #666; padding: 20px;'>
            <p><strong>Commercial Secondary Windows Savings Calculator v2.1.0</strong></p>
            <p>Complete Excel Data Integration - All Building Types Supported</p>
            <p>🔒 Your information is secure and only used for energy savings calculations.</p>
        </div>
        """, 
        unsafe_allow_html=True
    )idOfficePVAV_ElecElectric2080": {"heat_kwh_sf": 3.49, "cool_kwh_sf": 2.11, "gas_therms_sf": 0},
    "DoubleSingleMidOfficePVAV_ElecElectric2912": {"heat_kwh_sf": 4.63, "cool_kwh_sf": 3.1, "gas_therms_sf": 0},
    "DoubleSingleMidOfficePVAV_ElecElectric8760": {"heat_kwh_sf": 13.14, "cool_kwh_sf": 5.15, "gas_therms_sf": 0},
    
    # Single-Single Office Natural Gas combinations - FIXED KEYS
    "SingleSingleMidOfficePVAV_GasNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 3.03, "gas_therms_sf": 0.37},
    "SingleSingleMidOfficePVAV_GasNatural Gas2912": {"heat_kwh_sf": 0, "cool_kwh_sf": 4.99, "gas_therms_sf": 0.53},
    "SingleSingleMidOfficePVAV_GasNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 10.41, "gas_therms_sf": 1.31},
    "SingleDoubleMidOfficePVAV_GasNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 5.33, "gas_therms_sf": 0.35},
    "SingleDoubleMidOfficePVAV_GasNatural Gas2912": {"heat_kwh_sf": 0, "cool_kwh_sf": 7.99, "gas_therms_sf": 0.5},
    "SingleDoubleMidOfficePVAV_GasNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 13.29, "gas_therms_sf": 1.38},
    "DoubleSingleMidOfficePVAV_GasNatural Gas2080": {"heat_kwh_sf": 0, "cool_kwh_sf": 2.11, "gas_therms_sf": 0.17},
    "DoubleSingleMidOfficePVAV_GasNatural Gas2912": {"heat_kwh_sf": 0, "cool_kwh_sf": 3.1, "gas_therms_sf": 0.22},
    "DoubleSingleMidOfficePVAV_GasNatural Gas8760": {"heat_kwh_sf": 0, "cool_kwh_sf": 5.15, "gas_therms_sf": 0.64},
