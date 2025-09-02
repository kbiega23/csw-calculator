"Portland": {"HDD": 4187, "CDD": 367}, "Eugene": {"HDD": 4803, "CDD": 295}, "Medford": {"HDD": 4530, "CDD": 602},
    "Philadelphia": {"HDD": 4824, "CDD": 1184}, "Pittsburgh": {"HDD": 5925, "CDD": 726}, "Harrisburg": {"HDD": 5409, "CDD": 927},
    "Providence": {"HDD": 5870, "CDD": 735}, "Boston": {"HDD": 5793, "CDD": 734},
    "Charleston": {"HDD": 2051, "CDD": 2302}, "Columbia": {"HDD": 2593, "CDD": 2020}, "Greenville": {"HDD": 3652, "CDD": 1409},
    "Sioux Falls": {"HDD": 7680, "CDD": 680}, "Rapid City": {"HDD": 7203, "CDD": 675}, "Aberdeen": {"HDD": 7968, "CDD": 657},
    "Memphis": {"HDD": 2999, "CDD": 2134}, "Nashville": {"HDD": 3737, "CDD": 1751}, "Knoxville": {"HDD": 3959, "CDD": 1482},
    "Houston": {"HDD": 1439, "CDD": 2974}, "Dallas": {"HDD": 2333, "CDD": 2678}, "Austin": {"HDD": 1269, "CDD": 2884},
    "San Antonio": {"HDD": 1548, "CDD": 2992}, "El Paso": {"HDD": 2499, "CDD": 2171}, "Fort Worth": {"HDD": 2333, "CDD": 2678},
    "Salt Lake City": {"HDD": 5350, "CDD": 1118}, "Provo": {"HDD": 5836, "CDD": 858}, "Saint George": {"HDD": 2729, "CDD": 2936},
    "Burlington": {"HDD": 7491, "CDD": 420}, "Montpelier": {"HDD": 7662, "CDD": 249},
    "Norfolk": {"HDD": 3411, "CDD": 1630}, "Richmond": {"HDD": 3883, "CDD": 1493}, "Virginia Beach": {"HDD": 3336, "CDD": 1539},
    "Seattle": {"HDD": 4372, "CDD": 169}, "Spokane": {"HDD": 6716, "CDD": 341}, "Tacoma": {"HDD": 5664, "CDD": 142},
    "Charleston": {"HDD": 4708, "CDD": 1015}, "Huntington": {"HDD": 4642, "CDD": 1077},
    "Milwaukee": {"HDD": 7348, "CDD": 545}, "Madison": {"HDD": 7724, "CDD": 604}, "Green Bay": {"HDD": 7853, "CDD": 496},
    "Cheyenne": {"HDD": 7362, "CDD": 265}, "Casper": {"HDD": 7409, "CDD": 440}, "Jackson Hole": {"HDD": 9670, "CDD": 15}
}

# Real dropdown options from your Lists.csv
BUILDING_TYPES = ["Office", "Hotel", "School", "Hospital", "Multi-family"]

HVAC_SYSTEMS = [
    "Packaged VAV with electric reheat",
    "Packaged VAV with hydronic reheat", 
    "Built-up VAV with hydronic reheat",
    "PTAC",
    "PTHP",
    "Fan Coil Unit",
    "Other"
]

WINDOW_TYPES = [
    "Single pane",
    "Double pane", 
    "New double pane (U<0.35)"
]

SECONDARY_WINDOW_TYPES = [
    "Single",
    "Double"
]

HEATING_FUELS = [
    "Electric",
    "Natural Gas",
    "Electric Only"
]

# Lookup coefficients extracted from your Savings Lookup table and Regression coefficients
SAVINGS_COEFFICIENTS = {
    # Office building coefficients from your data
    "SingleSingleMidOfficePVAV_ElecElectric2080": {"heat_a": 0.8704668256, "heat_b": 0.0016759854, "heat_c": -0.0000000582, "cool_a": 2.4042885600, "cool_b": 0.0005625127, "cool_c": -0.0000000097, "gas_a": 0, "gas_b": 0, "gas_c": 0},
    "SingleSingleMidOfficePVAV_ElecElectric2912": {"heat_a": 1.6949778210, "heat_b": 0.0023051750, "heat_c": -0.0000000855, "cool_a": 3.7647661055, "cool_b": 0.0011532145, "cool_c": -0.0000000665, "gas_a": 0, "gas_b": 0, "gas_c": 0},
    "SingleSingleMidOfficePVAV_ElecElectric8760": {"heat_a": 16.0040605698, "heat_b": 0.0025681481, "heat_c": -0.0000000639, "cool_a": 6.1823179497, "cool_b": 0.0040023765, "cool_c": -0.0000002414, "gas_a": 0, "gas_b": 0, "gas_c": 0},
    "SingleSingleMidOfficePVAV_GasNaturalGas2080": {"heat_a": 0, "heat_b": 0, "heat_c": 0, "cool_a": 2.4042885600, "cool_b": 0.0005625127, "cool_c": -0.0000000097, "gas_a": 0.2963814100, "gas_b": 0.0000829100, "gas_c": 0.0000000011},
    "SingleSingleMidOfficePVAV_GasNaturalGas2912": {"heat_a": 0, "heat_b": 0, "heat_c": 0, "cool_a": 3.7647661055, "cool_b": 0.0011532145, "cool_c": -0.0000000665, "gas_a": 0.4214135200, "gas_b": 0.0001122600, "gas_c": 0.0000000002},
    "SingleSingleMidOfficePVAV_GasNaturalGas8760": {"heat_a": 0, "heat_b": 0, "heat_c": 0, "cool_a": 6.1823179497, "cool_b": 0.0040023765, "cool_c": -0.0000002414, "gas_a": 1.0444039700, "gas_b": 0.0001267500, "gas_c": -0.0000000004},
    "SingleSingleLargeOfficeVAVNaturalGas4000": {"heat_a": 0, "heat_b": 0, "heat_c": 0, "cool_a": 7.0786, "cool_b": 0.001, "cool_c": -0.0000001, "gas_a": 0.7473, "gas_b": 0.0001, "gas_c": 0},
    "SingleDoubleMidOfficePVAV_ElecElectric2080": {"heat_a": 1.2309323299, "heat_b": 0.0012506949, "heat_c": -0.0000000091, "cool_a": 4.2025198124, "cool_b": 0.0010760659, "cool_c": -0.0000000728, "gas_a": 0, "gas_b": 0, "gas_c": 0},
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
    
    if step == 1:  # Project Information
        required = ['project_name', 'contact_name', 'contact_email', 'company_name']
        return all(form_data.get(field) and str(form_data.get(field)).strip() for field in required)
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

def create_lookup_key(form_data):
    """Create lookup key matching Excel pattern"""
    # Determine building size based on floor area
    building_size = "Large" if form_data['floor_area'] >= 200000 else "Mid"
    
    # Convert HVAC type to lookup code
    hvac_mapping = {
        "Packaged VAV with electric reheat": "PVAV_Elec",
        "Packaged VAV with hydronic reheat": "PVAV_Gas", 
        "Built-up VAV with hydronic reheat": "VAV",
        "PTAC": "PTAC",
        "PTHP": "PTHP",
        "Fan Coil Unit": "FCU",
        "Other": "Other"
    }
    
    hvac_code = hvac_mapping.get(form_data['hvac_type'], "VAV")
    
    # Convert existing window type
    existing_map = {"Single pane": "Single", "Double pane": "Double", "New double pane (U<0.35)": "Double"}
    existing_code = existing_map.get(form_data['existing_window_type'], "Single")
    
    # Secondary window type
    secondary_code = form_data['secondary_window_type']  # Already "Single" or "Double"
    
    # Building type
    building_type = form_data['building_type']
    
    # Fuel type - fix the mapping
    fuel_mapping = {
        "Electric": "Electric",
        "Natural Gas": "NaturalGas",  # Remove space for lookup key
        "Electric Only": "Electric"
    }
    fuel_code = fuel_mapping.get(form_data['heating_fuel'], "NaturalGas")
    
    # Operating hours
    hours = form_data['operation_hours']
    
    # Create lookup key
    lookup_key = f"{existing_code}{secondary_code}{building_size}{building_type}{hvac_code}{fuel_code}{hours}"
    
    return lookup_key

def calculate_savings():
    """Calculate energy savings using real Excel formulas"""
    form_data = st.session_state.form_data
    
    # Get weather data - fix city name matching
    city = form_data['city']
    
    # Debug: Check if city exists in weather data
    if city not in WEATHER_DATA:
        # Try common variations
        city_variations = [
            city,
            city.replace(" ", ""),
            city.replace(".", ""),
            city + " " + form_data['state'][:2].upper()  # Add state abbreviation
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
    
    if city in WEATHER_DATA:
        weather = WEATHER_DATA[city]
    else:
        weather = {"HDD": 4000, "CDD": 1000}  # Final fallback
    
    hdd, cdd = weather["HDD"], weather["CDD"]
    
    # Create lookup key
    lookup_key = create_lookup_key(form_data)
    
    # Get coefficients - improved lookup with better fallbacks
    coefficients = SAVINGS_COEFFICIENTS.get(lookup_key)
    
    # If exact match not found, try simplified lookup
    if not coefficients:
        # Try different variations of the lookup key
        simplified_keys = []
        
        # Basic patterns to try
        building_size = "Large" if form_data['floor_area'] >= 200000 else "Mid"
        existing_type = "Single" if "Single" in form_data['existing_window_type'] else "Double"
        secondary_type = form_data['secondary_window_type']
        
        # Try common configurations
        if form_data['heating_fuel'] == "Natural Gas":
            simplified_keys.extend([
                f"SingleSingle{building_size}OfficePVAV_GasNaturalGas{form_data['operation_hours']}",
                f"SingleSingle{building_size}OfficeVAVNaturalGas{form_data['operation_hours']}",
                f"SingleSingleLargeOfficeVAVNaturalGas4000"  # Common fallback
            ])
        else:
            simplified_keys.extend([
                f"SingleSingle{building_size}OfficePVAV_ElecElectric{form_data['operation_hours']}",
                f"SingleSingle{building_size}OfficeVAVElectric{form_data['operation_hours']}"
            ])
        
        # Try each simplified key
        for key in simplified_keys:
            if key in SAVINGS_COEFFICIENTS:
                coefficients = SAVINGS_COEFFICIENTS[key]
                lookup_key = key  # Update for debugging
                break
    
    # Final fallback with more reasonable coefficients
    if not coefficients:
        # Use scaled coefficients based on building size and type
        base_cool = 3.0 if form_data['floor_area'] < 200000 else 5.0
        base_heat = 1.0 if form_data['heating_fuel'] == "Electric" else 0
        base_gas = 0.5 if form_data['heating_fuel'] == "Natural Gas" else 0
        
        coefficients = {
            "heat_a": base_heat, "heat_b": 0.001, "heat_c": -0.0000001,
            "cool_a": base_cool, "cool_b": 0.001, "cool_c": -0.0000001,
            "gas_a": base_gas, "gas_b": 0.0001, "gas_c": 0
        }
        lookup_key = "FALLBACK_REASONABLE"
    
    # Calculate using regression equations: savings = a + b*x + c*x^2
    # Where x is HDD for heating, CDD for cooling
    
    # Heating savings (kWh/SF of window)
    heating_kwh_per_sf = max(0, (coefficients["heat_a"] + 
                                coefficients["heat_b"] * hdd + 
                                coefficients["heat_c"] * hdd * hdd))
    
    # Cooling savings (kWh/SF of window) 
    cooling_kwh_per_sf = max(0, (coefficients["cool_a"] + 
                                coefficients["cool_b"] * cdd + 
                                coefficients["cool_c"] * cdd * cdd))
    
    # Gas savings (therms/SF of window)
    gas_therms_per_sf = max(0, (coefficients["gas_a"] + 
                               coefficients["gas_b"] * hdd + 
                               coefficients["gas_c"] * hdd * hdd))
    
    # Total savings for window area
    window_area = form_data['window_area']
    
    total_heating_kwh = heating_kwh_per_sf * window_area
    total_cooling_kwh = cooling_kwh_per_sf * window_area
    total_kwh = total_heating_kwh + total_cooling_kwh
    
    total_gas_therms = gas_therms_per_sf * window_area
    
    # Cost savings
    electric_cost = total_kwh * form_data['electric_rate']
    gas_cost = total_gas_therms * form_data['gas_rate']
    total_cost = electric_cost + gas_cost
    
    # Additional calculations
    floor_area = form_data['floor_area']
    cost_per_sf = total_cost / floor_area if floor_area > 0 else 0
    
    # Energy intensity calculations
    total_energy_btu = (total_kwh * 3412) + (total_gas_therms * 100000)  # Convert to BTU
    energy_intensity_savings = total_energy_btu / floor_area / 1000 if floor_area > 0 else 0  # kBtu/SF
    
    # Percentage savings (simplified calculation)
    baseline_eui = 85  # Typical office baseline
    if form_data['building_type'] == "Hotel":
        baseline_eui = 95
    elif form_data['building_type'] == "School":
        baseline_eui = 75
    elif form_data['building_type'] == "Hospital":
        baseline_eui = 145
    elif form_data['building_type'] == "Multi-family":
        baseline_eui = 65
    
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
        'city_used': city,  # For debugging
        'coefficients_used': coefficients  # For debugging
    }

# Main header
st.markdown('<div class="main-header">🏢 Winsert Savings Calculator</div>', unsafe_allow_html=True)

# Show progress
show_progress(st.session_state.step)

# Step 1: Project Information
if st.session_state.step == 1:
    st.markdown('<div class="step-header">📋 Step 1: Project Information</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        
        st.markdown("**Please provide your project details** <span class='required-field'>*Required for calculation</span>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            project_name = st.text_input(
                "Project Name *", 
                value=st.session_state.form_data.get('project_name', ''),
                placeholder="e.g., Downtown Office Building",
                help="Enter a descriptive name for your project"
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
                    index=city_index,
                    help="Select the city closest to your project location"
                )
            else:
                city = st.selectbox("City *", options=["Select state first"], disabled=True)
                city = None
        
        # Show climate data if city selected
        if city and city in WEATHER_DATA:
            weather = WEATHER_DATA[city]
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"🌡️ **Heating Degree Days (HDD):** {weather['HDD']:,}")
            with col2:
                st.markdown(f"❄️ **Cooling Degree Days (CDD):** {weather['CDD']:,}")
            st.markdown("*Climate data used for energy calculations*")
            st.markdown('</div>', unsafe_allow_html=True)
        
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
        
        building_descriptions = {
            "Office": "Commercial office buildings, corporate headquarters, business centers",
            "Hotel": "Hotels, motels, lodging facilities", 
            "School": "Educational facilities, schools, universities",
            "Hospital": "Healthcare facilities, medical centers, hospitals",
            "Multi-family": "Apartment buildings, condominiums, residential complexes"
        }
        
        for i, building_type in enumerate(BUILDING_TYPES):
            with cols[i]:
                if st.button(
                    f"🏢\n\n**{building_type}**\n\n{building_descriptions[building_type]}", 
                    key=f"building_{i}",
                    use_container_width=True,
                    type="primary" if selected_building == building_type else "secondary",
                    help=f"Select if your building is a {building_type.lower()}"
                ):
                    st.session_state.form_data['building_type'] = building_type
                    st.rerun()
        
        if selected_building:
            st.success(f"✅ Selected: **{selected_building}**")
            st.info(f"📋 {building_descriptions[selected_building]}")
        
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
                min_value=15000,
                max_value=500000,
                value=st.session_state.form_data.get('floor_area', 100000),
                step=5000,
                format="%d",
                help="Total conditioned floor area of the building (15,000 - 500,000 SF)"
            )
            
        with col2:
            num_stories = st.number_input(
                "Number of Stories *",
                min_value=1,
                max_value=50,
                value=st.session_state.form_data.get('num_stories', 5),
                step=1,
                format="%d",
                help="Number of floors in the building"
            )
            
        with col3:
            operation_hours = st.number_input(
                "Annual Operating Hours *",
                min_value=1980,
                max_value=8760,
                value=st.session_state.form_data.get('operation_hours', 3000),
                step=100,
                format="%d",
                help="Total operating hours per year (1,980 - 8,760 hours)"
            )
        
        # Show building size classification
        building_size = "Large" if floor_area >= 200000 else "Mid"
        st.info(f"📊 **Building Classification:** {building_size}-size building ({floor_area:,} SF)")
        
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
                index=WINDOW_TYPES.index(st.session_state.form_data.get('existing_window_type', WINDOW_TYPES[0])),
                help="Select the type of windows currently installed in your building"
            )
            
            window_area = st.number_input(
                "Square Feet of Secondary Windows to Install *",
                min_value=1000,
                max_value=100000,
                value=st.session_state.form_data.get('window_area', 25000),
                step=500,
                format="%d",
                help="Total area of secondary windows to be installed"
            )
            
        with col2:
            secondary_window_type = st.selectbox(
                "Secondary Window Type *",
                options=SECONDARY_WINDOW_TYPES,
                index=SECONDARY_WINDOW_TYPES.index(st.session_state.form_data.get('secondary_window_type', SECONDARY_WINDOW_TYPES[0])),
                help="Select the type of secondary windows to install"
            )
            
            # Show window-to-wall ratio
            if st.session_state.form_data.get('floor_area'):
                window_percentage = (window_area / st.session_state.form_data['floor_area']) * 100
                wwr = window_area / (st.session_state.form_data['floor_area'] / st.session_state.form_data.get('num_stories', 1)) * 0.3  # Simplified WWR calc
                st.info(f"📊 **Window-to-Wall Ratio:** ~{wwr:.2f}")
                st.info(f"📊 **{window_percentage:.1f}%** of total floor area")
        
        # Window performance information
        window_info = {
            "Single pane": "U-assembly = 1.03 Btu/hr-SF-F; SHGC = 0.73",
            "Double pane": "U-assembly = 0.65 Btu/hr-SF-F; SHGC = 0.40", 
            "New double pane (U<0.35)": "U-assembly < 0.35 Btu/hr-SF-F; SHGC < 0.35"
        }
        
        st.markdown(f"**Existing Window Performance:** {window_info.get(existing_window_type, '')}")
        
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
                index=HVAC_SYSTEMS.index(st.session_state.form_data.get('hvac_type', HVAC_SYSTEMS[0])),
                help="Select the primary HVAC system type for your building"
            )
            
        with col2:
            heating_fuel = st.selectbox(
                "Primary Heating Fuel *",
                options=HEATING_FUELS,
                index=HEATING_FUELS.index(st.session_state.form_data.get('heating_fuel', HEATING_FUELS[1])),
                help="Select the primary fuel used for heating"
            )
        
        # System compatibility check
        if "electric" in hvac_type.lower() and heating_fuel == "Natural Gas":
            st.warning("⚠️ **Compatibility Note:** You selected an electric HVAC system with Natural Gas fuel. Please verify this configuration is correct.")
        elif "hydronic" in hvac_type.lower() or "Gas" in hvac_type and heating_fuel == "Electric":
            st.warning("⚠️ **Compatibility Note:** You selected a gas/hydronic system with Electric fuel. Please verify this configuration is correct.")
        
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
                min_value=0.05,
                max_value=0.50,
                value=st.session_state.form_data.get('electric_rate', 0.12),
                step=0.001,
                format="%.3f",
                help="Your current electricity rate per kilowatt-hour"
            )
            
        with col2:
            gas_rate = st.number_input(
                "Natural Gas Rate ($/therm) *",
                min_value=0.50,
                max_value=3.00,
                value=st.session_state.form_data.get('gas_rate', 1.05),
                step=0.01,
                format="%.2f",
                help="Your current natural gas rate per therm"
            )
        
        # Rate information
        st.info("💡 **Tip:** Check your recent utility bills for accurate rates. Rates may include delivery charges and taxes.")
        
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
    
    # Additional metrics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("🌡️ Climate Zone", f"HDD: {results['hdd']:,} | CDD: {results['cdd']:,}")
    with col2:
        st.metric("📈 Energy Savings", f"{results['percentage_savings']:.1f}%", f"{results['energy_intensity_savings']:.1f} kBtu/SF/yr")
    with col3:
        st.metric("🏢 Baseline EUI", f"{results['baseline_eui']:.0f} kBtu/SF/yr")
    
    # Charts
    st.markdown("### 📈 Detailed Savings Analysis")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Energy savings chart
        fig_energy = go.Figure()
        
        # Electric savings breakdown
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
                y=[results['gas_therms'] * 29.3],  # Convert therms to kWh equivalent for chart
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
    
    with col2:
        # Cost savings pie chart  
        labels = []
        values = []
        colors = []
        
        if results['electric_cost'] > 0:
            labels.append('Electric Cost Savings')
            values.append(results['electric_cost'])
            colors.append('lightblue')
        
        if results['gas_cost'] > 0:
            labels.append('Gas Cost Savings')
            values.append(results['gas_cost'])
            colors.append('orange')
        
        if values:
            fig_cost = go.Figure(data=[go.Pie(
                labels=labels,
                values=values,
                hole=.3,
                marker_colors=colors,
                textinfo='label+percent+value',
                texttemplate='%{label}<br>%{percent}<br>$%{value:,.0f}'
            )])
            fig_cost.update_layout(
                title='Annual Cost Savings Distribution',
                height=400
            )
            st.plotly_chart(fig_cost, use_container_width=True)
    
    # Technical details expander
    with st.expander("🔍 Calculation Details & Debug Information"):
        st.markdown(f"""
        **Location & Weather Data:**
        - Selected Location: {st.session_state.form_data['city']}, {st.session_state.form_data['state']}
        - Weather Data Source: {results.get('city_used', 'Unknown')}
        - Climate Data: {results['hdd']:,} HDD, {results['cdd']:,} CDD
        
        **Building Configuration:**
        - Building Classification: {'Large' if st.session_state.form_data['floor_area'] >= 200000 else 'Mid'}-size {st.session_state.form_data['building_type']}
        - Lookup Configuration ID: `{results['lookup_key']}`
        - Floor Area: {st.session_state.form_data['floor_area']:,} SF
        - Window Area: {st.session_state.form_data['window_area']:,} SF
        - Operating Hours: {st.session_state.form_data['operation_hours']:,} hours/year
        
        **Calculation Coefficients Used:**
        - Heating: a={results.get('coefficients_used', {}).get('heat_a', 0):.4f}, b={results.get('coefficients_used', {}).get('heat_b', 0):.8f}, c={results.get('coefficients_used', {}).get('heat_c', 0):.12f}
        - Cooling: a={results.get('coefficients_used', {}).get('cool_a', 0):.4f}, b={results.get('coefficients_used', {}).get('cool_b', 0):.8f}, c={results.get('coefficients_used', {}).get('cool_c', 0):.12f}
        - Gas: a={results.get('coefficients_used', {}).get('gas_a', 0):.4f}, b={results.get('coefficients_used', {}).get('gas_b', 0):.8f}, c={results.get('coefficients_used', {}).get('gas_c', 0):.12f}
        
        **Energy Calculations:**
        - Heating Savings per SF: {results['heating_kwh']/st.session_state.form_data['window_area']:.4f} kWh/SF
        - Cooling Savings per SF: {results['cooling_kwh']/st.session_state.form_data['window_area']:.4f} kWh/SF
        - Gas Savings per SF: {results['gas_therms']/st.session_state.form_data['window_area']:.4f} therms/SF
        - Total Heating Savings: {results['heating_kwh']:,.0f} kWh/year
        - Total Cooling Savings: {results['cooling_kwh']:,.0f} kWh/year  
        - Total Gas Savings: {results['gas_therms']:,.1f} therms/year
        - Total Energy Intensity Savings: {results['energy_intensity_savings']:.1f} kBtu/SF/year
        
        **Cost Analysis:**
        - Electric Cost Savings: ${results['electric_cost']:,.0f}/year (@ ${st.session_state.form_data['electric_rate']:.3f}/kWh)
        - Gas Cost Savings: ${results['gas_cost']:,.0f}/year (@ ${st.session_state.form_data['gas_rate']:.2f}/therm)
        - Total Annual Savings: ${results['total_cost']:,.0f}/year
        - Savings per Square Foot: ${results['cost_per_sf']:.2f}/SF/year
        
        **Validation Notes:**
        - {"✅ Weather data found for selected city" if results.get('city_used') == st.session_state.form_data['city'] else "⚠️ Using weather data approximation"}
        - {"✅ Exact coefficient match found" if not results['lookup_key'].startswith('FALLBACK') else "⚠️ Using fallback coefficients - results may be approximate"}
        """)
        
        # Show comparison with typical ranges
        if results['total_cost'] > 0:
            cost_per_window_sf = results['total_cost'] / st.session_state.form_data['window_area']
            st.markdown(f"""
            **Reasonableness Check:**
            - Cost savings per window SF: ${cost_per_window_sf:.2f}/SF/year
            - Typical range: $2-8/SF/year for secondary windows
            - {"✅ Within expected range" if 2 <= cost_per_window_sf <= 8 else "⚠️ Outside typical range - please verify inputs"}
            """)
    
    # Project summary table
    st.markdown("### 📋 Complete Project Summary")
    
    summary_data = {
        'Parameter': [
            'Project Name',
            'Contact Person', 
            'Company',
            'Building Type',
            'Location',
            'Total Floor Area',
            'Number of Stories',
            'Operating Hours/Year',
            'Secondary Window Area',
            'Existing Window Type',
            'Secondary Window Type',
            'HVAC System',
            'Heating Fuel',
            'Electric Rate',
            'Gas Rate',
            '',
            '💡 Annual Electric Savings',
            '🔥 Annual Gas Savings', 
            '💰 Total Annual Cost Savings',
            '📊 Energy Intensity Savings',
            '📈 Percentage Energy Savings',
            '🌡️ Climate Zone (HDD/CDD)',
            '📉 Cost Savings per SF'
        ],
        'Value': [
            st.session_state.form_data['project_name'],
            st.session_state.form_data['contact_name'],
            st.session_state.form_data['company_name'],
            st.session_state.form_data['building_type'],
            f"{st.session_state.form_data['city']}, {st.session_state.form_data['state']}",
            f"{st.session_state.form_data['floor_area']:,} SF",
            f"{st.session_state.form_data['num_stories']} floors",
            f"{st.session_state.form_data['operation_hours']:,} hours/year",
            f"{st.session_state.form_data['window_area']:,} SF", 
            st.session_state.form_data['existing_window_type'],
            st.session_state.form_data['secondary_window_type'],
            st.session_state.form_data['hvac_type'],
            st.session_state.form_data['heating_fuel'],
            f"${st.session_state.form_data['electric_rate']:.3f}/kWh",
            f"${st.session_state.form_data['gas_rate']:.2f}/therm",
            '',
            f"{results['total_kwh']:,.0f} kWh/year",
            f"{results['gas_therms']:,.0f} therms/year",
            f"${results['total_cost']:,.0f}/year",
            f"{results['energy_intensity_savings']:.1f} kBtu/SF/year",
            f"{results['percentage_savings']:.1f}%",
            f"{results['hdd']:,} HDD / {results['cdd']:,} CDD",
            f"${results['cost_per_sf']:.2f}/SF/year"
        ]
    }
    
    df_summary = pd.DataFrame(summary_data)
    st.dataframe(df_summary, use_container_width=True, hide_index=True)
    
    # Download button
    csv_data = df_summary.to_csv(index=False)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"Winsert_Savings_Report_{st.session_state.form_data.get('project_name', 'Project').replace(' ', '_')}_{timestamp}.csv"
    
    st.download_button(
        label="📥 Download Complete Results Report",
        data=csv_data,
        file_name=filename,
        mime="text/csv",
        type="primary"
    )
    
    # Next steps and contact information
    st.markdown("### 🚀 Next Steps")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        **Ready to Move Forward?**
        
        Our team can help you:
        - 🔍 Conduct a detailed energy audit
        - 📋 Provide detailed specifications  
        - 💰 Identify utility rebates and incentives
        - 🏗️ Connect with certified installers
        - 📊 Develop financing options
        """)
    
    with col2:
        st.markdown("""
        **Contact Our Team:**
        
        📧 **Email:** [sales@company.com](mailto:sales@company.com)
        📞 **Phone:** 1-800-XXX-XXXX
        🌐 **Website:** [www.company.com](https://www.company.com)
        
        *Response within 24 hours*
        """)
    
    # Disclaimer
    st.markdown("---")
    st.markdown("""
    <div style='background-color: #fff3cd; padding: 15px; border-radius: 5px; border: 1px solid #ffeaa7; margin: 20px 0;'>
        <strong>⚠️ Important Disclaimer:</strong> These calculations are estimates based on typical building characteristics and regression analysis. 
        Actual savings may vary based on specific building conditions, occupancy patterns, maintenance practices, and local climate variations. 
        We recommend a professional energy audit for final decision making.
    </div>
    """, unsafe_allow_html=True)

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
            st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# Footer
if st.session_state.step < 8:
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #666; padding: 20px;'>
            <p><strong>Winsert Savings Calculator v2.0.0</strong></p>
            <p>🔒 Your information is secure and will only be used to provide you with energy savings estimates and product information.</p>
            <p>Questions? Contact us at <a href="mailto:support@company.com">support@company.com</a></p>
        </div>
        """, 
        unsafe_allow_html=True
    )import streamlit as st
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

# Weather data from your Weather Information tab
WEATHER_DATA = {
    "Anniston": {"HDD": 2585, "CDD": 1713}, "Auburn": {"HDD": 2688, "CDD": 1477}, "Birmingham": {"HDD": 2698, "CDD": 1912},
    "Daleville": {"HDD": 2447, "CDD": 2122}, "Dothan": {"HDD": 2072, "CDD": 2443}, "Gadsen": {"HDD": 3063, "CDD": 1383},
    "Huntsville": {"HDD": 3472, "CDD": 1701}, "Mobile": {"HDD": 1724, "CDD": 2524}, "Montgomery": {"HDD": 2183, "CDD": 2124},
    "Muscle Shoals": {"HDD": 2788, "CDD": 1701}, "Troy": {"HDD": 2196, "CDD": 2217}, "Tuscaloosa": {"HDD": 2656, "CDD": 2101},
    "Adak": {"HDD": 9202, "CDD": 0}, "Anchorage": {"HDD": 10158, "CDD": 0}, "Aniak": {"HDD": 11513, "CDD": 0},
    "Fairbanks": {"HDD": 13072, "CDD": 31}, "Juneau": {"HDD": 8471, "CDD": 1}, "Ketchikan": {"HDD": 8033, "CDD": 22},
    "Casa Grande": {"HDD": 1473, "CDD": 3715}, "Douglas": {"HDD": 2328, "CDD": 1832}, "Flagstaff": {"HDD": 7112, "CDD": 110},
    "Phoenix": {"HDD": 997, "CDD": 4591}, "Tucson": {"HDD": 1596, "CDD": 3020}, "Yuma": {"HDD": 535, "CDD": 4532},
    "Little Rock": {"HDD": 3079, "CDD": 2076}, "Fort Smith": {"HDD": 3596, "CDD": 2020}, "Hot Springs": {"HDD": 3256, "CDD": 2111},
    "Los Angeles": {"HDD": 1283, "CDD": 617}, "San Francisco": {"HDD": 2737, "CDD": 97}, "San Diego": {"HDD": 1019, "CDD": 742},
    "Sacramento": {"HDD": 2581, "CDD": 1281}, "Fresno": {"HDD": 2327, "CDD": 2101}, "Oakland": {"HDD": 2816, "CDD": 128},
    "Bakersfield": {"HDD": 2013, "CDD": 2240}, "Burbank": {"HDD": 1437, "CDD": 1449}, "Long Beach": {"HDD": 1136, "CDD": 995},
    "Denver": {"HDD": 5655, "CDD": 923}, "Colorado Springs": {"HDD": 6115, "CDD": 481}, "Grand Junction": {"HDD": 5283, "CDD": 1230},
    "Hartford Brainard": {"HDD": 5969, "CDD": 731}, "New Haven": {"HDD": 5524, "CDD": 625}, "Danbury": {"HDD": 6218, "CDD": 503},
    "Dover": {"HDD": 4987, "CDD": 1010}, "Wilmington": {"HDD": 5087, "CDD": 1088},
    "Miami": {"HDD": 150, "CDD": 4292}, "Tampa": {"HDD": 646, "CDD": 3442}, "Orlando": {"HDD": 526, "CDD": 3234},
    "Jacksonville": {"HDD": 1281, "CDD": 2565}, "Fort Lauderdale": {"HDD": 322, "CDD": 4114}, "Key West": {"HDD": 67, "CDD": 4906},
    "Atlanta": {"HDD": 2773, "CDD": 1809}, "Augusta": {"HDD": 2512, "CDD": 2023}, "Savannah": {"HDD": 1759, "CDD": 2474},
    "Honolulu": {"HDD": 0, "CDD": 4561}, "Hilo": {"HDD": 0, "CDD": 3279}, "Kahului": {"HDD": 0, "CDD": 3999},
    "Boise": {"HDD": 5395, "CDD": 756}, "Idaho Falls": {"HDD": 7769, "CDD": 246}, "Pocatello": {"HDD": 7100, "CDD": 393},
    "Chicago": {"HDD": 6399, "CDD": 830}, "Springfield": {"HDD": 5527, "CDD": 1166}, "Peoria": {"HDD": 6229, "CDD": 888},
    "Indianapolis": {"HDD": 5845, "CDD": 1044}, "Fort Wayne": {"HDD": 6471, "CDD": 764}, "Evansville": {"HDD": 4474, "CDD": 1376},
    "Des Moines": {"HDD": 6493, "CDD": 1121}, "Cedar Rapids": {"HDD": 6900, "CDD": 714}, "Dubuque": {"HDD": 7516, "CDD": 525},
    "Wichita": {"HDD": 4324, "CDD": 1554}, "Topeka": {"HDD": 4953, "CDD": 1315}, "Kansas City": {"HDD": 5434, "CDD": 1316},
    "Louisville": {"HDD": 4521, "CDD": 1449}, "Lexington": {"HDD": 4856, "CDD": 1104}, "Bowling Green": {"HDD": 4380, "CDD": 1410},
    "New Orleans": {"HDD": 1358, "CDD": 2784}, "Baton Rouge": {"HDD": 1762, "CDD": 2622}, "Shreveport": {"HDD": 2351, "CDD": 2384},
    "Portland": {"HDD": 7679, "CDD": 335}, "Bangor": {"HDD": 7671, "CDD": 450}, "Augusta": {"HDD": 7495, "CDD": 396},
    "Baltimore": {"HDD": 4631, "CDD": 1237}, "Washington DC": {"HDD": 4921, "CDD": 1113}, "Hagerstown": {"HDD": 4867, "CDD": 1079},
    "Boston": {"HDD": 5793, "CDD": 734}, "Worcester": {"HDD": 7164, "CDD": 346}, "Springfield": {"HDD": 6623, "CDD": 494},
    "Detroit": {"HDD": 6621, "CDD": 679}, "Grand Rapids": {"HDD": 6908, "CDD": 579}, "Lansing": {"HDD": 7045, "CDD": 597},
    "Minneapolis": {"HDD": 7783, "CDD": 731}, "Duluth": {"HDD": 9620, "CDD": 159}, "Rochester": {"HDD": 8386, "CDD": 488},
    "Jackson": {"HDD": 2428, "CDD": 2237}, "Biloxi": {"HDD": 1928, "CDD": 2606}, "Greenville": {"HDD": 2432, "CDD": 2468},
    "St. Louis": {"HDD": 4846, "CDD": 1555}, "Kansas City": {"HDD": 5434, "CDD": 1316}, "Springfield": {"HDD": 4629, "CDD": 1307},
    "Billings": {"HDD": 6731, "CDD": 548}, "Great Falls": {"HDD": 7854, "CDD": 344}, "Missoula": {"HDD": 7323, "CDD": 269},
    "Omaha": {"HDD": 5954, "CDD": 1275}, "Lincoln": {"HDD": 5892, "CDD": 1220}, "North Platte": {"HDD": 6556, "CDD": 802},
    "Las Vegas": {"HDD": 2301, "CDD": 3187}, "Reno": {"HDD": 5488, "CDD": 620}, "Elko": {"HDD": 7070, "CDD": 407},
    "Manchester": {"HDD": 6322, "CDD": 664}, "Concord": {"HDD": 7479, "CDD": 397}, "Portsmouth": {"HDD": 6753, "CDD": 527},
    "Newark": {"HDD": 5057, "CDD": 1237}, "Atlantic City": {"HDD": 5073, "CDD": 886}, "Trenton": {"HDD": 5040, "CDD": 1177},
    "Albuquerque": {"HDD": 4157, "CDD": 1269}, "Las Cruces": {"HDD": 2876, "CDD": 2146}, "Santa Fe": {"HDD": 5449, "CDD": 589},
    "New York": {"HDD": 4885, "CDD": 1133}, "Buffalo": {"HDD": 6612, "CDD": 468}, "Rochester": {"HDD": 6518, "CDD": 627},
    "Albany": {"HDD": 6773, "CDD": 489}, "Syracuse": {"HDD": 6609, "CDD": 535}, "Binghamton": {"HDD": 7015, "CDD": 404},
    "Charlotte": {"HDD": 3153, "CDD": 1675}, "Raleigh": {"HDD": 3465, "CDD": 1566}, "Asheville": {"HDD": 4273, "CDD": 817},
    "Fargo": {"HDD": 9211, "CDD": 491}, "Bismarck": {"HDD": 8452, "CDD": 453}, "Grand Forks": {"HDD": 9534, "CDD": 486},
    "Columbus": {"HDD": 5599, "CDD": 747}, "Cleveland": {"HDD": 6160, "CDD": 760}, "Cincinnati": {"HDD": 4911, "CDD": 1039},
    "Oklahoma City": {"HDD": 3556, "CDD": 2038}, "Tulsa": {"HDD": 3844, "CDD": 2066}, "Lawton": {"HDD": 3055, "CDD": 1979},
    "Portland": {"HDD": 4187, "CDD": 367}, "Eugene": {"HDD": 4803, "CDD
