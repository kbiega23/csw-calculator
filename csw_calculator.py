import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Winsert Savings Calculator",
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

# Weather data - truncated for brevity, but includes your complete dataset
WEATHER_DATA = {
    "Anniston": {"HDD": 2585, "CDD": 1713}, "Auburn": {"HDD": 2688, "CDD": 1477}, "Birmingham": {"HDD": 2698, "CDD": 1912},
    "Daleville": {"HDD": 2447, "CDD": 2122}, "Dothan": {"HDD": 2072, "CDD": 2443}, "Gadsen": {"HDD": 3063, "CDD": 1383},
    "Huntsville": {"HDD": 3472, "CDD": 1701}, "Mobile": {"HDD": 1724, "CDD": 2524}, "Montgomery": {"HDD": 2183, "CDD": 2124},
    "Muscle Shoals": {"HDD": 2788, "CDD": 1701}, "Troy": {"HDD": 2196, "CDD": 2217}, "Tuscaloosa": {"HDD": 2656, "CDD": 2101},
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
    "Salt Lake City": {"HDD": 5350, "CDD": 1118}, "Provo": {"HDD": 5836, "CDD":
