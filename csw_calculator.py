import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Commercial Secondary Windows Savings Calculator",
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
    .navigation-buttons {
        display: flex;
        justify-content: space-between;
        margin: 30px 0;
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
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'form_data' not in st.session_state:
    st.session_state.form_data = {}

# Sample data - Replace with actual data from your Excel sheets
STATES_CITIES = {
    "Alabama": ["Birmingham", "Mobile", "Montgomery", "Huntsville", "Tuscaloosa"],
    "Alaska": ["Anchorage", "Fairbanks", "Juneau"],
    "Arizona": ["Phoenix", "Tucson", "Mesa", "Scottsdale", "Flagstaff"],
    "California": ["Los Angeles", "San Francisco", "San Diego", "Sacramento", "Fresno", "Oakland", "San Jose"],
    "Colorado": ["Denver", "Colorado Springs", "Aurora", "Fort Collins", "Boulder"],
    "Connecticut": ["Hartford", "New Haven", "Stamford", "Waterbury", "Norwalk"],
    "Florida": ["Miami", "Tampa", "Orlando", "Jacksonville", "Fort Lauderdale", "Tallahassee"],
    "Georgia": ["Atlanta", "Augusta", "Columbus", "Savannah", "Athens"],
    "Illinois": ["Chicago", "Aurora", "Rockford", "Joliet", "Naperville", "Springfield"],
    "New York": ["New York City", "Buffalo", "Rochester", "Yonkers", "Syracuse", "Albany"],
    "Texas": ["Houston", "Dallas", "Austin", "San Antonio", "Fort Worth", "El Paso"],
    "Washington": ["Seattle", "Spokane", "Tacoma", "Vancouver", "Bellevue"]
    # Add more states and cities from your Excel data
}

WEATHER_DATA = {
    "New York City": {"HDD": 4885, "CDD": 1133},
    "Buffalo": {"HDD": 6927, "CDD": 465},
    "Rochester": {"HDD": 6748, "CDD": 515},
    "Chicago": {"HDD": 6455, "CDD": 865},
    "Los Angeles": {"HDD": 1405, "CDD": 1060},
    "San Francisco": {"HDD": 3015, "CDD": 109},
    "Miami": {"HDD": 130, "CDD": 4459},
    "Denver": {"HDD": 5810, "CDD": 650},
    "Seattle": {"HDD": 4640, "CDD": 145},
    "Phoenix": {"HDD": 974, "CDD": 3400},
    "Atlanta": {"HDD": 2826, "CDD": 1667}
    # Add more cities from your Weather Information tab
}

BUILDING_TYPES = ["Office", "Hotel", "School", "Hospital", "Multi-family"]

HVAC_SYSTEMS = [
    "Central Air Conditioning with Gas Heat",
    "Central Air Conditioning with Electric Heat", 
    "Heat Pump",
    "Package Terminal Heat Pump (PTHP)",
    "Window Units",
    "Split System"
]

WINDOW_TYPES = [
    "Single Pane Clear",
    "Single Pane Tinted", 
    "Dual Pane Clear",
    "Dual Pane Tinted",
    "Dual Pane Low-E"
]

SECONDARY_WINDOW_TYPES = [
    "Single Pane Secondary",
    "Dual Pane Secondary"
]

HEATING_FUELS = [
    "Natural Gas",
    "Electricity", 
    "Oil",
    "Propane"
]

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
    elif direction == "prev":
        st.session_state.step -= 1
    
    st.experimental_rerun()

def validate_current_step():
    """Validate required fields for current step"""
    step = st.session_state.step
    form_data = st.session_state.form_data
    
    if step == 1:  # Project Information
        required = ['project_name', 'contact_name', 'contact_email', 'company_name']
        return all(form_data.get(field) for field in required)
    elif step == 2:  # Location
        required = ['state', 'city']
        return all(form_data.get(field) for field in required)
    elif step == 3:  # Building Type
        return form_data.get('building_type') is not None
    elif step == 4:  # Building Parameters
        required = ['floor_area', 'num_stories', 'operation_hours']
        return all(form_data.get(field) for field in required)
    elif step == 5:  # Window Specifications  
        required = ['existing_window_type', 'secondary_window_type', 'window_area']
        return all(form_data.get(field) for field in required)
    elif step == 6:  # HVAC System
        required = ['hvac_type', 'heating_fuel']
        return all(form_data.get(field) for field in required)
    elif step == 7:  # Utility Rates
        required = ['electric_rate', 'gas_rate']
        return all(form_data.get(field) for field in required)
    
    return True

def calculate_savings():
    """Calculate energy savings based on form data"""
    form_data = st.session_state.form_data
    
    # Get weather data
    city = form_data['city']
    weather = WEATHER_DATA.get(city, {"HDD": 4000, "CDD": 1000})
    hdd, cdd = weather["HDD"], weather["CDD"]
    
    # Simplified calculation - replace with actual regression equations
    building_type = form_data['building_type']
    floor_area = form_data['floor_area']
    window_area = form_data['window_area']
    
    # Base energy intensity by building type (kBtu/SF/year)
    base_intensity = {
        "Office": 85, "Hotel": 95, "School": 75, 
        "Hospital": 145, "Multi-family": 65
    }
    
    # Window improvement factors
    window_factors = {
        ("Single Pane Clear", "Dual Pane Secondary"): 1.8,
        ("Single Pane Tinted", "Dual Pane Secondary"): 1.7,
        ("Dual Pane Clear", "Dual Pane Secondary"): 1.3,
        ("Dual Pane Tinted", "Dual Pane Secondary"): 1.2,
        ("Dual Pane Low-E", "Dual Pane Secondary"): 1.1
    }
    
    window_key = (form_data['existing_window_type'], form_data['secondary_window_type'])
    window_factor = window_factors.get(window_key, 1.2)
    
    # Climate factors
    heating_factor = min(hdd / 4000, 2.0)
    cooling_factor = min(cdd / 1000, 2.0)
    
    # Calculate savings
    base_eff = base_intensity[building_type]
    
    # Energy savings per SF of secondary windows
    heating_savings_per_sf = 15.0 * window_factor * heating_factor
    cooling_savings_per_sf = 12.0 * window_factor * cooling_factor
    
    # Total savings
    total_heating_kwh = heating_savings_per_sf * window_area * 0.293  # kBtu to kWh
    total_cooling_kwh = cooling_savings_per_sf * window_area * 0.293
    
    # Gas savings for gas heating
    gas_therms = 0
    if form_data['heating_fuel'] == "Natural Gas":
        gas_therms = heating_savings_per_sf * window_area * 0.01
        total_heating_kwh = total_heating_kwh * 0.2  # Reduce electric if gas heating
    
    total_kwh = total_heating_kwh + total_cooling_kwh
    
    # Cost savings
    electric_cost = total_kwh * form_data['electric_rate']
    gas_cost = gas_therms * form_data['gas_rate']
    total_cost = electric_cost + gas_cost
    
    return {
        'heating_kwh': total_heating_kwh,
        'cooling_kwh': total_cooling_kwh,
        'total_kwh': total_kwh,
        'gas_therms': gas_therms,
        'electric_cost': electric_cost,
        'gas_cost': gas_cost,
        'total_cost': total_cost,
        'cost_per_sf': total_cost / floor_area,
        'hdd': hdd,
        'cdd': cdd
    }

# Main header
st.markdown('<div class="main-header">🏢 Commercial Secondary Windows Savings Calculator</div>', unsafe_allow_html=True)

# Show progress
show_progress(st.session_state.step)

# Step 1: Project Information
if st.session_state.step == 1:
    st.markdown('<div class="step-header">📋 Step 1: Project Information</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Please provide your project details** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            project_name = st.text_input(
                "Project Name *", 
                value=st.session_state.form_data.get('project_name', ''),
                placeholder="e.g., Downtown Office Building"
            )
            contact_name = st.text_input(
                "Your Name *", 
                value=st.session_state.form_data.get('contact_name', ''),
                placeholder="First Last"
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
            placeholder="123 Main Street, City, State"
        )
        
        phone = st.text_input(
            "Phone Number", 
            value=st.session_state.form_data.get('phone', ''),
            placeholder="(555) 123-4567"
        )
        
        # Save data
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
    st.markdown('<div class="step-header">🌍 Step 2: Select Your Location</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Select your project location for climate data** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
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
        
        # Show climate data if city selected
        if city and city in WEATHER_DATA:
            weather = WEATHER_DATA[city]
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"🌡️ **Heating Degree Days:** {weather['HDD']:,}")
            with col2:
                st.info(f"❄️ **Cooling Degree Days:** {weather['CDD']:,}")
        
        # Save data
        st.session_state.form_data.update({
            'state': state if state else None,
            'city': city if city else None
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 3: Building Type
elif st.session_state.step == 3:
    st.markdown('<div class="step-header">🏢 Step 3: Select Building Type</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**What type of building is this project?** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
        # Building type cards
        cols = st.columns(len(BUILDING_TYPES))
        selected_building = st.session_state.form_data.get('building_type')
        
        for i, building_type in enumerate(BUILDING_TYPES):
            with cols[i]:
                if st.button(
                    f"🏢\n\n**{building_type}**", 
                    key=f"building_{i}",
                    use_container_width=True,
                    type="primary" if selected_building == building_type else "secondary"
                ):
                    st.session_state.form_data['building_type'] = building_type
                    st.experimental_rerun()
        
        if selected_building:
            st.success(f"✅ Selected: **{selected_building}**")
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 4: Building Parameters
elif st.session_state.step == 4:
    st.markdown('<div class="step-header">📐 Step 4: Building Parameters</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Enter your building specifications** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            floor_area = st.number_input(
                "Total Floor Area (SF) *",
                min_value=1000,
                max_value=10000000,
                value=st.session_state.form_data.get('floor_area', 50000),
                step=1000,
                format="%d"
            )
            
        with col2:
            num_stories = st.number_input(
                "Number of Stories *",
                min_value=1,
                max_value=100,
                value=st.session_state.form_data.get('num_stories', 5),
                step=1,
                format="%d"
            )
            
        with col3:
            operation_hours = st.number_input(
                "Annual Operating Hours *",
                min_value=1000,
                max_value=8760,
                value=st.session_state.form_data.get('operation_hours', 2080),
                step=100,
                format="%d"
            )
        
        # Save data
        st.session_state.form_data.update({
            'floor_area': floor_area,
            'num_stories': num_stories,
            'operation_hours': operation_hours
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 5: Window Specifications
elif st.session_state.step == 5:
    st.markdown('<div class="step-header">🪟 Step 5: Window Specifications</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Specify your window configuration** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            existing_window_type = st.selectbox(
                "Existing Window Type *",
                options=WINDOW_TYPES,
                index=WINDOW_TYPES.index(st.session_state.form_data.get('existing_window_type', WINDOW_TYPES[0]))
            )
            
            window_area = st.number_input(
                "Square Feet of Secondary Windows to Install *",
                min_value=100,
                max_value=1000000,
                value=st.session_state.form_data.get('window_area', 15000),
                step=100,
                format="%d"
            )
            
        with col2:
            secondary_window_type = st.selectbox(
                "Secondary Window Type *",
                options=SECONDARY_WINDOW_TYPES,
                index=SECONDARY_WINDOW_TYPES.index(st.session_state.form_data.get('secondary_window_type', SECONDARY_WINDOW_TYPES[0]))
            )
            
            # Show percentage of floor area
            if st.session_state.form_data.get('floor_area'):
                window_percentage = (window_area / st.session_state.form_data['floor_area']) * 100
                st.info(f"📊 **{window_percentage:.1f}%** of total floor area")
        
        # Save data
        st.session_state.form_data.update({
            'existing_window_type': existing_window_type,
            'secondary_window_type': secondary_window_type,
            'window_area': window_area
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 6: HVAC System
elif st.session_state.step == 6:
    st.markdown('<div class="step-header">🔧 Step 6: HVAC System</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Select your HVAC configuration** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
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
                index=HEATING_FUELS.index(st.session_state.form_data.get('heating_fuel', HEATING_FUELS[0]))
            )
        
        # Save data
        st.session_state.form_data.update({
            'hvac_type': hvac_type,
            'heating_fuel': heating_fuel
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 7: Utility Rates
elif st.session_state.step == 7:
    st.markdown('<div class="step-header">💰 Step 7: Utility Rates</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Enter your current utility rates** <span class='required-field'>*Required</span>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            electric_rate = st.number_input(
                "Electric Rate ($/kWh) *",
                min_value=0.01,
                max_value=1.0,
                value=st.session_state.form_data.get('electric_rate', 0.12),
                step=0.001,
                format="%.3f"
            )
            
        with col2:
            gas_rate = st.number_input(
                "Natural Gas Rate ($/therm) *",
                min_value=0.01,
                max_value=5.0,
                value=st.session_state.form_data.get('gas_rate', 1.05),
                step=0.01,
                format="%.2f"
            )
        
        # Save data
        st.session_state.form_data.update({
            'electric_rate': electric_rate,
            'gas_rate': gas_rate
        })
        
        st.markdown('</div>', unsafe_allow_html=True)

# Step 8: Results
elif st.session_state.step == 8:
    st.markdown('<div class="step-header">📊 Step 8: Your Energy Savings Results</div>', unsafe_allow_html=True)
    
    # Calculate results
    results = calculate_savings()
    
    # Results display
    st.markdown("### 🎉 Congratulations! Here are your projected energy savings:")
    
    # Key metrics cards
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="result-card">
            <h3>💡 Electric Savings</h3>
            <h2>{results['total_kwh']:,.0f}</h2>
            <p>kWh per year</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="result-card">
            <h3>🔥 Gas Savings</h3>
            <h2>{results['gas_therms']:,.0f}</h2>
            <p>therms per year</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="result-card">
            <h3>💰 Annual Savings</h3>
            <h2>${results['total_cost']:,.0f}</h2>
            <p>per year</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="result-card">
            <h3>📊 Cost per SF</h3>
            <h2>${results['cost_per_sf']:.2f}</h2>
            <p>per square foot</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Charts
    st.markdown("### 📈 Savings Breakdown")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Energy savings chart
        fig_energy = go.Figure(data=[
            go.Bar(name='Heating', x=['Electric (kWh)', 'Gas (therms)'], 
                  y=[results['heating_kwh'], results['gas_therms']], 
                  marker_color='orange'),
            go.Bar(name='Cooling', x=['Electric (kWh)', 'Gas (therms)'], 
                  y=[results['cooling_kwh'], 0], 
                  marker_color='lightblue')
        ])
        fig_energy.update_layout(title='Annual Energy Savings', barmode='stack')
        st.plotly_chart(fig_energy, use_container_width=True)
    
    with col2:
        # Cost savings pie chart  
        fig_cost = go.Figure(data=[go.Pie(
            labels=['Electric Cost Savings', 'Gas Cost Savings'],
            values=[results['electric_cost'], results['gas_cost']],
            marker_colors=['lightblue', 'orange']
        )])
        fig_cost.update_layout(title='Annual Cost Savings Distribution')
        st.plotly_chart(fig_cost, use_container_width=True)
    
    # Project summary table
    st.markdown("### 📋 Project Summary")
    
    summary_data = {
        'Parameter': [
            'Project Name', 'Building Type', 'Location', 'Floor Area (SF)',
            'Window Area (SF)', 'HVAC System', 'Electric Rate', 'Gas Rate',
            '', 'Annual Electric Savings', 'Annual Gas Savings', 
            'Total Annual Cost Savings', 'Savings per Square Foot'
        ],
        'Value': [
            st.session_state.form_data['project_name'],
            st.session_state.form_data['building_type'],
            f"{st.session_state.form_data['city']}, {st.session_state.form_data['state']}",
            f"{st.session_state.form_data['floor_area']:,}",
            f"{st.session_state.form_data['window_area']:,}",
            st.session_state.form_data['hvac_type'],
            f"${st.session_state.form_data['electric_rate']:.3f}/kWh",
            f"${st.session_state.form_data['gas_rate']:.2f}/therm",
            '',
            f"{results['total_kwh']:,.0f} kWh/year",
            f"{results['gas_therms']:,.0f} therms/year",
            f"${results['total_cost']:,.0f}/year",
            f"${results['cost_per_sf']:.2f}/SF/year"
        ]
    }
    
    df_summary = pd.DataFrame(summary_data)
    st.dataframe(df_summary, use_container_width=True, hide_index=True)
    
    # Download button
    csv_data = df_summary.to_csv(index=False)
    st.download_button(
        label="📥 Download Results Report",
        data=csv_data,
        file_name=f"CSW_Savings_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
        type="primary"
    )
    
    # Contact information
    st.markdown("### 📞 Next Steps")
    st.info("""
    **Ready to move forward with your secondary windows project?**
    
    Our team is ready to help you:
    - Refine these calculations with a detailed energy audit
    - Provide product specifications and pricing
    - Connect you with certified installers in your area
    - Assist with utility rebate applications
    
    **Contact Information:**
    - Email: sales@company.com
    - Phone: 1-800-XXX-XXXX
    - Website: www.company.com
    """)

# Navigation buttons
st.markdown('<div class="navigation-buttons">', unsafe_allow_html=True)

col1, col2, col3 = st.columns([1, 2, 1])

with col1:
    if st.session_state.step > 1:
        if st.button("⬅️ Previous", use_container_width=True):
            navigate_step("prev")

with col3:
    if st.session_state.step < 8:
        if st.button("Next ➡️", use_container_width=True, type="primary"):
            navigate_step("next")

with col2:
    if st.session_state.step == 8:
        if st.button("🔄 Start New Calculation", use_container_width=True, type="secondary"):
            # Reset all session state
            st.session_state.step = 1
            st.session_state.form_data = {}
            st.experimental_rerun()

st.markdown('</div>', unsafe_allow_html=True)

# Footer
if st.session_state.step < 8:
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #666; padding: 20px;'>
            <p><strong>Commercial Secondary Windows Savings Calculator v2.0.0</strong></p>
            <p>🔒 Your information is secure and will only be used to provide you with energy savings estimates and product information.</p>
        </div>
        """, 
        unsafe_allow_html=True
    )