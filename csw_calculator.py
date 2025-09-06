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
            <h3>Gas Cost Savings</h3>
            <h2>${results['gas_cost']:,.0f}</h2>
            <p>per year</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Key metrics (Excel format)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Savings", f"${results['total_cost']:,.0f}/yr", f"{results['energy_intensity_savings']:.1f} kBtu/SF-yr")
    with col2:
        st.metric("Baseline Energy Use Intensity", f"{results['baseline_eui']:.1f} kBtu/SF-year")
    with col3:
        st.metric("Savings Percentage", f"{results['percentage_savings']:.1f}%", f"Est. WWR: {results['wwr']:.2f}")
    
    # Detailed breakdown (Excel-style)
    st.markdown("### Detailed Energy Savings Analysis")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Per SF calculations (Excel methodology)
        window_area = st.session_state.form_data['window_area']
        
        st.markdown("**Energy Savings per SF of CSW:**")
        
        breakdown_data = {
            'Category': [
                'Heating, kWh/SFCSW',
                'Cooling & Fans, kWh/SFCSW', 
                'Heating, therms/SFCSW',
                'Total Electric, kWh/SFCSW',
                'Electric Cost, $/SFCSW/yr',
                'Gas Cost, $/SFCSW/yr'
            ],
            'Value': [
                f"{results['heating_kwh']/window_area:.2f}",
                f"{results['cooling_kwh']/window_area:.2f}",
                f"{results['gas_therms']/window_area:.2f}",
                f"{results['total_kwh']/window_area:.2f}",
                f"${results['electric_cost']/window_area:.2f}",
                f"${results['gas_cost']/window_area:.2f}"
            ]
        }
        
        df_breakdown = pd.DataFrame(breakdown_data)
        st.dataframe(df_breakdown, use_container_width=True, hide_index=True)
    
    with col2:
        # Charts
        fig_energy = go.Figure()
        
        if results['heating_kwh'] > 0:
            fig_energy.add_trace(go.Bar(
                name='Heating (Electric)',
                x=['Annual Savings'],
                y=[results['heating_kwh']],
                marker_color='orange',
                text=f"{results['heating_kwh']:,.0f} kWh",
                textposition='inside'
            ))
        
        fig_energy.add_trace(go.Bar(
            name='Cooling (Electric)', 
            x=['Annual Savings'],
            y=[results['cooling_kwh']],
            marker_color='lightblue',
            text=f"{results['cooling_kwh']:,.0f} kWh",
            textposition='inside'
        ))
        
        if results['gas_therms'] > 0:
            fig_energy.add_trace(go.Bar(
                name='Gas Savings',
                x=['Annual Savings'],
                y=[results['gas_therms'] * 29.3],  # Convert to kWh equivalent for display
                marker_color='green',
                text=f"{results['gas_therms']:,.0f} therms",
                textposition='inside'
            ))
        
        fig_energy.update_layout(
            title='Annual Energy Savings Breakdown',
            yaxis_title='Energy Savings (kWh equivalent)',
            barmode='stack',
            height=400,
            showlegend=True
        )
        st.plotly_chart(fig_energy, use_container_width=True)
    
    # Excel-style calculation details
    with st.expander("Calculation Details & Lookup Information"):
        st.markdown(f"""
        **Building Configuration:**
        - Building Classification: {'Large' if st.session_state.form_data['floor_area'] >= 200000 else 'Mid'}-size {st.session_state.form_data['building_type']}
        - Lookup Key Used: `{results['lookup_key']}`
        - Floor Area: {st.session_state.form_data['floor_area']:,} SF
        - CSW Area: {st.session_state.form_data['window_area']:,} SF
        - Operating Hours: {st.session_state.form_data['operation_hours']:,} hours/year
        - Estimated Wall Area: {results['wall_area']:,.0f} SF
        - Window-to-Wall Ratio: {results['wwr']:.3f} ({results['wwr']*100:.1f}%)
        
        **Climate Data:**
        - Location: {st.session_state.form_data['city']}, {st.session_state.form_data['state']}
        - Weather Data Source: {results.get('city_used', 'Unknown')}
        - Heating Degree Days (HDD): {results['hdd']:,}
        - Cooling Degree Days (CDD): {results['cdd']:,}
        
        **Savings Coefficients (from Excel Lookup Table):**
        - Heating Savings: {results['savings_data_used']['heat_kwh_sf']:.2f} kWh/SFCSW
        - Cooling Savings: {results['savings_data_used']['cool_kwh_sf']:.2f} kWh/SFCSW
        - Gas Savings: {results['savings_data_used']['gas_therms_sf']:.2f} therms/SFCSW
        
        **Cost Analysis:**
        - Electric Rate: ${st.session_state.form_data['electric_rate']:.3f}/kWh
        - Gas Rate: ${st.session_state.form_data['gas_rate']:.2f}/therm
        - Total Annual Savings: ${results['total_cost']:,.0f}/year
        - Savings per Building SF: ${results['cost_per_sf']:.2f}/SF/year
        
        **Energy Intensity:**
        - Total Energy Savings: {results['energy_intensity_savings']:.1f} kBtu/SF/year
        - Baseline EUI: {results['baseline_eui']:.1f} kBtu/SF/year
        - Percentage Energy Savings: {results['percentage_savings']:.1f}%
        """)
        
        # WWR validation (Excel-style)
        if results['wwr'] < 0.1:
            st.warning("WWR is below typical range (0.1-0.5). Results may not be representative of typical installations.")
        elif results['wwr'] > 0.5:
            st.warning("WWR is above typical range (0.1-0.5). Please verify window area input.")
    
    # Complete project summary (Excel format)
    st.markdown("### Complete Project Summary")
    
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
            'Gas Cost Savings, $/yr',
            'Total Savings, $/yr',
            'Total Savings, kBtu/SF-yr',
            'Baseline Energy Use Intensity, kBtu/SF-year',
            'Savings, % and kBtu/SF-year'
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
            f"$ {st.session_state.form_data['electric_rate']:.2f}",
            f"$ {st.session_state.form_data['gas_rate']:.2f}",
            f"{results['wwr']:.2f}",
            '',
            f"{results['total_kwh']:,.0f}",
            f"{results['gas_therms']:,.0f}",
            f"$ {results['electric_cost']:,.0f}",
            f"$ {results['gas_cost']:,.0f}",
            f"$ {results['total_cost']:,.0f}",
            f"{results['energy_intensity_savings']:.1f}",
            f"{results['baseline_eui']:.1f}",
            f"{results['percentage_savings']:.1f}% and {results['energy_intensity_savings']:.1f}"
        ]
    }
    
    df_summary = pd.DataFrame(summary_data)
    st.dataframe(df_summary, use_container_width=True, hide_index=True)
    
    # Download functionality
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
    
    # Disclaimer (Excel-style)
    st.markdown("---")
    st.markdown("""
    <div style='background-color: #fff3cd; padding: 15px; border-radius: 5px; border: 1px solid #ffeaa7; margin: 20px 0;'>
        <strong>Important Disclaimer:</strong> These calculations are estimates based on building energy modeling and regression analysis. 
        Actual savings may vary based on specific building conditions, occupancy patterns, maintenance practices, and local climate variations. 
        We recommend a professional energy audit for final decision making.
    </div>
    """, unsafe_allow_html=True)

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
    )import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import math

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

# Weather data from Excel
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
    "Salt Lake City": {"HDD": 5350, "CDD": 1118}, "Provo": {"HDD": 5836, "CDD": 858}, "Saint George": {"HDD": 2729, "CDD": 2936},
    "Burlington": {"HDD": 7491, "CDD": 420}, "Montpelier": {"HDD": 7662, "CDD": 249},
    "Norfolk": {"HDD": 3411, "CDD": 1630}, "Richmond": {"HDD": 3883, "CDD": 1493}, "Virginia Beach": {"HDD": 3336, "
