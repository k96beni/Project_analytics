import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import tempfile
import os
import time
from datetime import datetime
import calendar
import base64
from PIL import Image as PILImage
import io

# Change

# Set page configuration
st.set_page_config(
    page_title="ChargeNode Rapport Generator",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        border-radius: 4px 4px 0px 0px;
    }
    h1, h2, h3 {
        color: #2c3e50;
    }
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 20px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        text-align: center;
        height: 150px; /* Ensure consistent height */
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .metric-value {
        font-size: 2.5rem;
        font-weight: bold;
        color: #27ae60;
    }
    .metric-label {
        font-size: 1.2rem;
        color: #7f8c8d;
    }
</style>
""", unsafe_allow_html=True)

# Page title
st.title("⚡ ChargeNode Rapport Generator - Internt verktyg för rapportskapande")
st.markdown("Ladda upp dina datafiler för att analysera prestanda och användning av ladduttag.")

# File upload section
st.sidebar.header("Ladda upp Datafiler")
sessions_file = st.sidebar.file_uploader("Ladda upp Sessions.xlsx", type=["xlsx"])
overview_file = st.sidebar.file_uploader("Ladda upp Overview.xlsx (Valfri)", type=["xlsx"]) # Made optional

# Expected column names and their potential alternatives
EXPECTED_COLUMNS = {
    'Startad': ['Startad', 'Start', 'Started', 'Start Time', 'StartTime', 'Beginning', 'Start Date'],
    'Avslutad': ['Avslutad', 'End', 'Ended', 'End Time', 'EndTime', 'Finish', 'End Date'],
    'Laddat (kWh)': ['Laddat (kWh)', 'kWh', 'Energy', 'Charged', 'Energy (kWh)', 'Power', 'Consumption'],
    'Omsättning (exkl)': ['Kostnad (exkl)', 'Cost', 'Price', 'Cost (excl)', 'Fee', 'Charge', 'Amount'],
    'Uttag': ['Uttag', 'Outlet', 'Charger', 'Outlet ID', 'Charger ID', 'Terminal', 'Station'],
    'Område': ['Område', 'Area', 'Location', 'Region', 'Zone', 'Site', 'Place']
}

# Function to find matching columns
def find_matching_columns(df, expected_columns_dict):
    column_mapping = {}
    available_columns = df.columns.tolist()
    
    for expected_col, alternatives in expected_columns_dict.items():
        found = False
        # First try exact matches
        for alt in alternatives:
            if alt in available_columns:
                column_mapping[expected_col] = alt # Store as expected:actual
                found = True
                break
        
        # If no exact match, try case-insensitive matching
        if not found:
            for col in available_columns:
                for alt in alternatives:
                    if col.lower() == alt.lower():
                        column_mapping[expected_col] = col # Store as expected:actual
                        found = True
                        break
                if found:
                    break
    
    # Invert mapping for renaming: actual_name -> expected_name
    rename_map = {v: k for k, v in column_mapping.items()}
    return rename_map

# Function to extract area code and name from combined format
def extract_area_code_and_name(area_string):
    """
    Extract area code and name from a string in the format '2571 - Stena Hildedalsgatan'
    Returns tuple of (code, name)
    """
    try:
        if not isinstance(area_string, str):
            return (str(area_string), str(area_string))
            
        area_string = area_string.strip()
        if ' - ' in area_string:
            # Format is "2571 - Stena Hildedalsgatan"
            parts = area_string.split(' - ', 1)
            code = parts[0].strip()
            name = parts[1].strip() if len(parts) > 1 else ""
            return (code, name)
        else:
            # Just return the whole string if it doesn't match the expected format
            return (area_string, area_string)
    except:
        return (str(area_string), str(area_string))

# Function to preprocess data
def preprocess_data(sessions_df, overview_df):
    st.write("Tillgängliga kolumner i Sessions-filen:", sessions_df.columns.tolist())
    if overview_df is not None:
        st.write("Tillgängliga kolumner i Overview-filen:", overview_df.columns.tolist())
    
    column_rename_map = find_matching_columns(sessions_df, EXPECTED_COLUMNS)
    if column_rename_map:
        sessions_df = sessions_df.rename(columns=column_rename_map)
    
    required_columns = ['Startad', 'Avslutad', 'Laddat (kWh)', 'Omsättning (exkl)', 'Uttag', 'Område']
    missing_columns = [col for col in required_columns if col not in sessions_df.columns]
    
    if missing_columns:
        st.warning(f"Följande obligatoriska kolumner saknas: {', '.join(missing_columns)}")
        st.warning("Skapar platshållardata för saknade kolumner. För korrekt analys, se till att din data har dessa kolumner.")
        
        if 'Startad' not in sessions_df.columns: sessions_df['Startad'] = pd.to_datetime('2023-01-01T00:00:00')
        if 'Avslutad' not in sessions_df.columns: sessions_df['Avslutad'] = pd.to_datetime('2023-01-01T01:00:00')
        if 'Laddat (kWh)' not in sessions_df.columns: sessions_df['Laddat (kWh)'] = 10.0
        if 'Omsättning (exkl)' not in sessions_df.columns: sessions_df['Omsättning (exkl)'] = 5.0
        if 'Uttag' not in sessions_df.columns: sessions_df['Uttag'] = "Uttag_1"
        if 'Område' not in sessions_df.columns: sessions_df['Område'] = 'Default Area'
    
    try:
        sessions_df['Startad'] = pd.to_datetime(sessions_df['Startad'])
        sessions_df['Avslutad'] = pd.to_datetime(sessions_df['Avslutad'])
    except Exception as e:
        st.error(f"Fel vid konvertering av datumkolumner: {e}. Kontrollera formatet.")
        st.stop()
    
    sessions_df['Year'] = sessions_df['Startad'].dt.year
    sessions_df['Month'] = sessions_df['Startad'].dt.month
    sessions_df['Month_Name'] = sessions_df['Startad'].dt.strftime('%b')
    sessions_df['Year_Month'] = sessions_df['Startad'].dt.strftime('%Y-%m')
    sessions_df['Hour'] = sessions_df['Startad'].dt.hour
    sessions_df['Weekday'] = sessions_df['Startad'].dt.day_name()
    sessions_df['Day_Type'] = np.where(sessions_df['Startad'].dt.weekday < 5, 'Vardag', 'Helg')

    sessions_df['Duration_Minutes'] = (sessions_df['Avslutad'] - sessions_df['Startad']).dt.total_seconds() / 60
    sessions_df['Duration_Hours'] = sessions_df['Duration_Minutes'] / 60
    sessions_df.loc[sessions_df['Duration_Hours'] < 0, 'Duration_Hours'] = 0
    sessions_df.loc[sessions_df['Duration_Minutes'] < 0, 'Duration_Minutes'] = 0

    for col in ['Laddat (kWh)', 'Omsättning (exkl)']:
        if col in sessions_df.columns:
            if sessions_df[col].dtype == object:
                sessions_df[col] = (
                    sessions_df[col]
                    .astype(str)
                    .str.replace('\xa0', '', regex=False)  # ta bort icke-brytande mellanslag
                    .str.replace(' ', '', regex=False)     # ta bort vanliga mellanslag (för säkerhets skull)
                    .str.replace(',', '.', regex=False)    # omvandlar svenska decimalkomma till punkt
                )
            sessions_df[col] = pd.to_numeric(sessions_df[col], errors='coerce').fillna(0)
    
    if 'Uttag' in sessions_df.columns:
        sessions_df['Uttag'] = sessions_df['Uttag'].astype(str)

    # Remove sessions with zero or negative duration
    sessions_df = sessions_df[sessions_df['Duration_Minutes'] > 1]
    
    # Handle area code formats in overview file
    if overview_df is not None and not overview_df.empty:
        first_col = overview_df.columns[0]
        
        # Check if the first column might contain area codes in the format "2571 - Stena Hildedalsgatan"
        sample_values = overview_df[first_col].astype(str).head().tolist()
        has_area_code_format = any(' - ' in val for val in sample_values)
        
        if has_area_code_format:
            # Extract area codes and names
            extracted = overview_df[first_col].apply(extract_area_code_and_name)
            overview_df['AreaCode'] = extracted.apply(lambda x: x[0])
            overview_df['AreaName'] = extracted.apply(lambda x: x[1])
            
            # Try to convert AreaCode to numeric if possible
            try:
                overview_df['AreaCode'] = pd.to_numeric(overview_df['AreaCode'], errors='coerce').fillna(overview_df['AreaCode'])
            except:
                pass  # Keep as string if conversion fails
        
        # Also check if there are separate area code and name columns in the sessions file
        if 'Område' in sessions_df.columns:
            # Look for potential area code column
            possible_code_columns = [col for col in sessions_df.columns if any(keyword in col.lower() for keyword in 
                                                   ['områdeskod', 'omradeskod', 'area code', 'areakod', 'kod'])]
            
            if possible_code_columns:
                area_code_col = possible_code_columns[0]
                
                # Ensure the area code column is properly formatted as string for comparison
                sessions_df[area_code_col] = sessions_df[area_code_col].astype(str).str.strip()
                
                # Look for a potential area name column
                possible_name_columns = [col for col in sessions_df.columns if any(keyword in col.lower() for keyword in 
                                                     ['områdesnamn', 'omradesnamn', 'area name', 'areanamn', 'namn'])]
                
                area_name_col = possible_name_columns[0] if possible_name_columns else None
                
                # If we have both area codes in the overview file and separate columns in sessions, we can match them
                if has_area_code_format:
                    # Create a mapping from area code to original format
                    area_mapping = {}
                    for _, row in overview_df.iterrows():
                        area_mapping[str(row['AreaCode'])] = row[first_col]
                    
                    # Update the Område column to match the format in overview file
                    sessions_df['Område_Original'] = sessions_df['Område']  # Keep original for reference
                    
                    # Update based on area code
                    sessions_df['Område'] = sessions_df[area_code_col].map(area_mapping).fillna(sessions_df['Område'])

    return sessions_df, overview_df

# Function to calculate metrics
def calculate_metrics(sessions_df, overview_df):
    metrics = {}
    
    if sessions_df.empty:
        st.warning("Ingen data att analysera efter filtrering eller i den uppladdade filen.")
        # Return default empty metrics
        metrics['unique_areas'] = []
        metrics['area_count'] = 0
        metrics['outlets_per_area'] = pd.DataFrame(columns=['Område', 'Number_of_Outlets'])
        metrics['total_outlets'] = 0
        metrics['kwh_per_month_area'] = pd.DataFrame(columns=['Område', 'Year_Month', 'Laddat (kWh)'])
        metrics['revenue_per_month_area'] = pd.DataFrame(columns=['Område', 'Year_Month', 'Omsättning (exkl)'])
        metrics['kwh_per_month'] = pd.DataFrame(columns=['Year_Month', 'Laddat (kWh)'])
        metrics['avg_kwh_per_outlet_month'] = pd.DataFrame(columns=['Year_Month', 'Avg_kWh_per_Outlet'])
        metrics['kwh_outlet_month_area'] = pd.DataFrame(columns=['Område', 'Year_Month', 'Uttag', 'Laddat (kWh)'])
        metrics['hourly_utilization'] = pd.DataFrame(columns=['Date', 'Hour', 'IsWeekend', 'Outlets_In_Use', 'Total_Outlets', 'Utilization', 'Date_Str'])
        metrics['utilization'] = pd.DataFrame(columns=['Område', 'Year_Month', 'Used_Outlet_Hours', 'Total_Possible_Outlet_Hours', 'Utilization'])
        metrics['total_kwh'] = 0
        metrics['total_revenue'] = 0
        metrics['total_sessions'] = 0
        metrics['avg_kwh_per_session'] = 0
        metrics['avg_duration_minutes'] = 0
        metrics['avg_hourly_utilization'] = pd.DataFrame(columns=['Hour', 'Day_Type', 'Utilization'])
        metrics['avg_weekday_utilization'] = pd.DataFrame(columns=['Weekday', 'Utilization'])
        return metrics

    unique_areas = sessions_df['Område'].unique()
    metrics['unique_areas'] = unique_areas
    metrics['area_count'] = len(unique_areas)
    
    # Calculate outlets per area from sessions
    outlets_per_area = sessions_df.groupby('Område')['Uttag'].nunique().reset_index()
    outlets_per_area.columns = ['Område', 'Number_of_Outlets']
    
    # If overview_df is available, use column 4 for a more accurate count of outlets
    total_outlets = 0
    if overview_df is not None and not overview_df.empty:
        try:
            # Use column 4 (index 3) for outlet counts as specified in the problem
            outlet_col_idx = 3  # 0-based indexing, so column 4 is index 3
            
            if outlet_col_idx < len(overview_df.columns):
                outlet_column = overview_df.columns[outlet_col_idx]
                
                # Use first column for area identification
                area_id_col = overview_df.columns[0]
                
                # If we extracted area codes earlier, use them for matching
                if 'AreaCode' in overview_df.columns:
                    # Convert column 4 to numeric if it's not already
                    if overview_df[outlet_column].dtype == object:
                        overview_df[outlet_column] = pd.to_numeric(overview_df[outlet_column], errors='coerce').fillna(0)
                    
                    # Group by area code and sum the outlet counts
                    outlet_counts = overview_df.groupby('AreaCode')[outlet_column].sum().reset_index()
                    outlet_counts.columns = ['AreaCode', 'Number_of_Outlets']
                    
                    # Create a mapping from original area format to area code for matching
                    area_code_map = {}
                    for _, row in overview_df.iterrows():
                        area_full = row[area_id_col]
                        area_code = row['AreaCode'] if 'AreaCode' in overview_df.columns else area_full
                        area_code_map[str(area_full)] = str(area_code)
                    
                    # Update outlets_per_area with counts from overview
                    updated_outlets_per_area = []
                    for _, row in outlets_per_area.iterrows():
                        area = row['Område']
                        num_outlets = row['Number_of_Outlets']
                        
                        # Try to find the area code from the session area
                        area_code = None
                        
                        # Check if the area is already in the format "code - name"
                        if ' - ' in str(area):
                            area_code = extract_area_code_and_name(area)[0]
                        else:
                            # Otherwise try to find it in the mapping
                            for full_name, code in area_code_map.items():
                                if str(area) == str(full_name) or str(area) == str(code):
                                    area_code = code
                                    break
                        
                        # If we found a matching area code, update the outlet count
                        if area_code is not None and str(area_code) in outlet_counts['AreaCode'].astype(str).values:
                            matching_row = outlet_counts[outlet_counts['AreaCode'].astype(str) == str(area_code)]
                            if not matching_row.empty:
                                num_outlets = matching_row.iloc[0]['Number_of_Outlets']
                        
                        updated_outlets_per_area.append({
                            'Område': area,
                            'Number_of_Outlets': num_outlets
                        })
                    
                    # Replace the original outlets_per_area with the updated one
                    outlets_per_area = pd.DataFrame(updated_outlets_per_area)
                else:
                    # If no AreaCode column, try matching directly by the first column
                    if overview_df[outlet_column].dtype == object:
                        overview_df[outlet_column] = pd.to_numeric(overview_df[outlet_column], errors='coerce').fillna(0)
                    
                    outlet_counts = overview_df.groupby(area_id_col)[outlet_column].sum().reset_index()
                    outlet_counts.columns = ['Område', 'Number_of_Outlets']
                    
                    # Try to update outlets_per_area based on area names
                    for i, row in outlets_per_area.iterrows():
                        area = row['Område']
                        for j, count_row in outlet_counts.iterrows():
                            overview_area = count_row['Område']
                            # Check for exact match or if one contains the other
                            if str(area) == str(overview_area) or str(area) in str(overview_area) or str(overview_area) in str(area):
                                outlets_per_area.at[i, 'Number_of_Outlets'] = count_row['Number_of_Outlets']
                                break
                
                # Calculate total from the updated outlets_per_area
                total_outlets = outlets_per_area['Number_of_Outlets'].sum()
            else:
                st.warning(f"Column 4 not found in overview file. Using data from sessions instead.")
                total_outlets = int(outlets_per_area['Number_of_Outlets'].sum())
        except Exception as e:
            st.warning(f"Kunde inte beräkna antalet uttag från översiktsfilen: {e}")
            import traceback
            st.warning(traceback.format_exc())  # More detailed error for debugging
            # Fall back to using only session data
            total_outlets = int(outlets_per_area['Number_of_Outlets'].sum())
    else:
        # If no overview file, just use the active outlets from sessions
        total_outlets = int(outlets_per_area['Number_of_Outlets'].sum())
    
    metrics['outlets_per_area'] = outlets_per_area
    metrics['total_outlets'] = total_outlets
    
    # For calculations that use total_outlets, make sure it's never zero to avoid division by zero
    total_outlets_for_calc = max(1, total_outlets)  # Use at least 1 to avoid division by zero

    kwh_per_month_area = sessions_df.groupby(['Område', 'Year_Month'])['Laddat (kWh)'].sum().reset_index()
    metrics['kwh_per_month_area'] = kwh_per_month_area

    revenue_per_month_area = sessions_df.groupby(['Område', 'Year_Month'])['Omsättning (exkl)'].sum().reset_index()
    metrics['revenue_per_month_area'] = revenue_per_month_area
    
    kwh_per_month = sessions_df.groupby(['Year_Month'])['Laddat (kWh)'].sum().reset_index()
    metrics['kwh_per_month'] = kwh_per_month
    
    total_kwh_per_month = sessions_df.groupby('Year_Month')['Laddat (kWh)'].sum().reset_index()
    outlets_active_per_month = sessions_df.groupby('Year_Month')['Uttag'].nunique().reset_index()
    avg_kwh_per_outlet = pd.merge(total_kwh_per_month, outlets_active_per_month, on='Year_Month', how='left')
    avg_kwh_per_outlet['Avg_kWh_per_Outlet'] = avg_kwh_per_outlet['Laddat (kWh)'] / avg_kwh_per_outlet['Uttag'].replace(0, 1) # Avoid div by zero
    metrics['avg_kwh_per_outlet_month'] = avg_kwh_per_outlet
    
    kwh_outlet_month_area = sessions_df.groupby(['Område', 'Year_Month', 'Uttag'])['Laddat (kWh)'].sum().reset_index()
    metrics['kwh_outlet_month_area'] = kwh_outlet_month_area
    
    # Calculate utilization by hour and date (Improved robustness)
    session_hours_list = []
    if not sessions_df.empty and 'Startad' in sessions_df.columns and 'Avslutad' in sessions_df.columns:
        for _, session in sessions_df.iterrows():
            try:
                start_time = session['Startad']
                end_time = session['Avslutad']
                outlet = session['Uttag']
                area = session['Område']
                
                if pd.isna(start_time) or pd.isna(end_time) or end_time < start_time:
                    continue
                
                # Iterate over one-hour intervals within the session
                current_hour_start = start_time.floor('H')
                while current_hour_start < end_time:
                    session_hours_list.append({
                        'Date': current_hour_start.date(),
                        'Hour': current_hour_start.hour,
                        'Outlet': outlet,
                        'Område': area,  # Add area for better segmentation
                        'IsWeekend': current_hour_start.weekday() >= 5,
                        'Day_Type': 'Helg' if current_hour_start.weekday() >= 5 else 'Vardag',
                        'Weekday': current_hour_start.day_name()
                    })
                    current_hour_start += pd.Timedelta(hours=1)
            except Exception as e:
                st.warning(f"Skippade en session vid beräkning av timvis användning p.g.a. datafel: {e}")
                continue
    
    if session_hours_list:
        hourly_usage_df = pd.DataFrame(session_hours_list)
        
        # Count unique outlets in use for each hour, date, day_type, weekday
        # First by area for area-specific utilization
        area_hourly_utilization = hourly_usage_df.groupby(['Date', 'Hour', 'IsWeekend', 'Day_Type', 'Weekday', 'Område']).agg(
            Outlets_In_Use=('Outlet', 'nunique')
        ).reset_index()
        
        # Add total configured outlets per area
        area_hourly_utilization = area_hourly_utilization.merge(
            outlets_per_area, on='Område', how='left'
        )
        area_hourly_utilization['Number_of_Outlets'] = area_hourly_utilization['Number_of_Outlets'].fillna(1)  # Fallback to 1 if unknown
        
        # Calculate area-specific utilization
        area_hourly_utilization['Utilization'] = area_hourly_utilization['Outlets_In_Use'] / area_hourly_utilization['Number_of_Outlets']
        area_hourly_utilization['Date_Str'] = pd.to_datetime(area_hourly_utilization['Date']).dt.strftime('%a %d %b %Y')
        metrics['area_hourly_utilization'] = area_hourly_utilization
        
        # Then overall for all areas combined
        hourly_utilization = hourly_usage_df.groupby(['Date', 'Hour', 'IsWeekend', 'Day_Type', 'Weekday']).agg(
            Outlets_In_Use=('Outlet', 'nunique')
        ).reset_index()
        
        # Use the total configured outlets for overall utilization
        hourly_utilization['Total_Outlets'] = total_outlets_for_calc
        hourly_utilization['Utilization'] = hourly_utilization['Outlets_In_Use'] / hourly_utilization['Total_Outlets']
        hourly_utilization['Date_Str'] = pd.to_datetime(hourly_utilization['Date']).dt.strftime('%a %d %b %Y')
        metrics['hourly_utilization'] = hourly_utilization

        # Average hourly utilization (e.g. 9 AM is X% utilized on average on weekdays)
        avg_hourly_utilization = hourly_utilization.groupby(['Hour', 'Day_Type'])['Utilization'].mean().reset_index()
        metrics['avg_hourly_utilization'] = avg_hourly_utilization

        # Average weekday utilization
        avg_weekday_utilization = hourly_utilization.groupby(['Weekday'])['Utilization'].mean().reset_index()
        weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        avg_weekday_utilization['Weekday'] = pd.Categorical(avg_weekday_utilization['Weekday'], categories=weekday_order, ordered=True)
        avg_weekday_utilization = avg_weekday_utilization.sort_values('Weekday')
        metrics['avg_weekday_utilization'] = avg_weekday_utilization

    else:
        metrics['hourly_utilization'] = pd.DataFrame(columns=['Date', 'Hour', 'IsWeekend', 'Outlets_In_Use', 'Total_Outlets', 'Utilization', 'Date_Str', 'Day_Type', 'Weekday'])
        metrics['area_hourly_utilization'] = pd.DataFrame(columns=['Date', 'Hour', 'IsWeekend', 'Outlets_In_Use', 'Number_of_Outlets', 'Utilization', 'Date_Str', 'Day_Type', 'Weekday', 'Område'])
        metrics['avg_hourly_utilization'] = pd.DataFrame(columns=['Hour', 'Day_Type', 'Utilization'])
        metrics['avg_weekday_utilization'] = pd.DataFrame(columns=['Weekday', 'Utilization'])

    # Monthly utilization
    utilization_data = []
    all_months_in_data = sessions_df['Year_Month'].unique()
    all_areas_in_data = sessions_df['Område'].unique()

    for area in all_areas_in_data:
        # Get total number of outlets for this area from our updated outlets_per_area dataframe
        area_outlets_count = outlets_per_area[outlets_per_area['Område'] == area]['Number_of_Outlets'].values
        if len(area_outlets_count) == 0 or area_outlets_count[0] == 0:
            continue  # Skip areas with no configured outlets
            
        area_outlets_count = area_outlets_count[0]

        for ym in all_months_in_data:
            year, month = map(int, ym.split('-'))
            days_in_month = calendar.monthrange(year, month)[1]
            total_possible_outlet_hours = area_outlets_count * days_in_month * 24 # Total available hours for all outlets in area for the month
            
            area_month_sessions = sessions_df[(sessions_df['Område'] == area) & (sessions_df['Year_Month'] == ym)]
            if area_month_sessions.empty:
                used_outlet_hours = 0
            else:
                # Calculate active hours from session duration
                # For each session, count the hours it spans
                total_hours = 0
                for _, session in area_month_sessions.iterrows():
                    try:
                        start_time = session['Startad']
                        end_time = session['Avslutad']
                        if pd.isna(start_time) or pd.isna(end_time) or end_time < start_time:
                            continue
                            
                        # Calculate hours this session spans
                        duration_hours = (end_time - start_time).total_seconds() / 3600
                        total_hours += duration_hours
                    except Exception as e:
                        st.warning(f"Fel vid beräkning av timanvändning för session: {e}")
                        continue
                
                used_outlet_hours = total_hours

            utilization = used_outlet_hours / total_possible_outlet_hours if total_possible_outlet_hours > 0 else 0
            
            utilization_data.append({
                'Område': area,
                'Year_Month': ym,
                'Used_Outlet_Hours': used_outlet_hours,
                'Total_Possible_Outlet_Hours': total_possible_outlet_hours,
                'Utilization': utilization
            })
    metrics['utilization'] = pd.DataFrame(utilization_data)
    
    metrics['total_kwh'] = sessions_df['Laddat (kWh)'].sum()
    metrics['total_revenue'] = sessions_df['Omsättning (exkl)'].sum()
    metrics['total_sessions'] = len(sessions_df)
    metrics['avg_kwh_per_session'] = metrics['total_kwh'] / metrics['total_sessions'] if metrics['total_sessions'] > 0 else 0
    metrics['avg_duration_minutes'] = sessions_df['Duration_Minutes'].mean() if metrics['total_sessions'] > 0 else 0
    
    return metrics

# Function to create plotly figures
def create_visualizations(metrics, sessions_df): # Pass sessions_df for more detailed plots
    figures = {}
    
    # 1. Bar chart: Number of outlets per area
    fig_outlets = px.bar(
        metrics['outlets_per_area'], x='Område', y='Number_of_Outlets',
        title='Antal Uttag per Område', color='Number_of_Outlets', color_continuous_scale=px.colors.sequential.Greens,
        labels={'Number_of_Outlets': 'Antal Uttag', 'Område': 'Område'}
    )
    fig_outlets.update_layout(xaxis_tickangle=-45)
    figures['outlets_per_area'] = fig_outlets
    
    # 2. Bar chart: kWh per month for each area
    fig_kwh_area = px.bar(
        metrics['kwh_per_month_area'], x='Year_Month', y='Laddat (kWh)', color='Område',
        title='Total Energi (kWh) per Månad och Område',
        labels={'Laddat (kWh)': 'Energi (kWh)', 'Year_Month': 'Månad', 'Område': 'Område'},
        barmode='group', color_discrete_sequence=px.colors.qualitative.Pastel
    )
    fig_kwh_area.update_layout(xaxis_tickangle=-45)
    figures['kwh_per_month_area'] = fig_kwh_area

    # 2b. Bar chart: Omsättning (exkl) per month for each area
    fig_revenue_area = px.bar(
        metrics['revenue_per_month_area'], x='Year_Month', y='Omsättning (exkl)', color='Område',
        title='Total Omsättning (exkl. moms) per Månad och Område',
        labels={'Omsättning (exkl)': 'Omsättning (SEK)', 'Year_Month': 'Månad', 'Område': 'Område'},
        barmode='group', color_discrete_sequence=px.colors.qualitative.Set2
    )
    fig_revenue_area.update_layout(xaxis_tickangle=-45)
    figures['revenue_per_month_area'] = fig_revenue_area
    
    # 3. Line chart: Average kWh per outlet per month (overall)
    fig_avg_kwh = px.line(
        metrics['avg_kwh_per_outlet_month'], x='Year_Month', y='Avg_kWh_per_Outlet', markers=True,
        title='Genomsnittlig Energi (kWh) per Aktivt Uttag per Månad (Totalt)',
        labels={'Avg_kWh_per_Outlet': 'Genomsnittlig kWh/Uttag', 'Year_Month': 'Månad'}
    )
    fig_avg_kwh.update_layout(xaxis_tickangle=-45)
    figures['avg_kwh_per_outlet'] = fig_avg_kwh
    
    # 4. Heatmap: Date-Hour Utilization - CHANGED TO GREEN GRADIENT
    if not metrics['hourly_utilization'].empty:
        pivot_hourly = metrics['hourly_utilization'].pivot_table(
            index='Date_Str', columns='Hour', values='Utilization', aggfunc='mean'
        ).fillna(0)
        # Ensure chronological order of dates
        sorted_dates = sorted(metrics['hourly_utilization']['Date'].unique())
        date_str_order = [d.strftime('%a %d %b %Y') for d in sorted_dates]
        pivot_hourly = pivot_hourly.reindex(index=date_str_order).fillna(0)
        
        # Find the min and max values for consistent color scale
        min_val = pivot_hourly.values.min()
        max_val = pivot_hourly.values.max()
        
        # Create the heatmap with green color scale
        fig_hourly = px.imshow(
            pivot_hourly, labels=dict(x='Timme på Dygnet', y='Datum', color='Beläggning (%)'),
            x=list(range(24)), y=pivot_hourly.index, 
            color_continuous_scale=px.colors.sequential.Greens,  # Changed to green
            title='Timvis Beläggningsgrad Heatmap (Aktiva Uttag / Totalt Antal Uttag)', 
            zmin=min_val, zmax=max_val  # Set fixed min/max for consistent gradient
        )
        
        # Add ChargeNode logo to top right (placeholder)
        fig_hourly.add_layout_image(
            dict(
                source="https://chargenode.eu/wp-content/themes/chargenode/dist/assets/img/logo.svg",
                xref="paper", yref="paper",
                x=1.0, y=1.05,
                sizex=0.15, sizey=0.15,
                xanchor="right", yanchor="top"
            )
        )
        
        fig_hourly.update_layout(
            coloraxis_colorbar=dict(title='Beläggning', tickformat='.0%'),
            height=max(500, 20 * len(pivot_hourly.index)), margin=dict(l=150, r=20, t=80, b=50),
            xaxis_title="Timme på dygnet", yaxis_title="Datum"
        )
        hour_labels = [f"{h:02d}:00" for h in range(24)]
        fig_hourly.update_xaxes(tickvals=list(range(24)), ticktext=hour_labels, side="top")
        fig_hourly.update_yaxes(autorange="reversed") # Show newest dates at top
        figures['hourly_utilization_heatmap'] = fig_hourly
    else:
        figures['hourly_utilization_heatmap'] = go.Figure().update_layout(title_text='Timvis Beläggningsgrad Heatmap (Ingen data)')

    # 4b. Aggregated Hourly Utilization (Line chart) - Changed to use green
    if not metrics['avg_hourly_utilization'].empty:
        fig_avg_hourly_util = px.line(
            metrics['avg_hourly_utilization'], x='Hour', y='Utilization', color='Day_Type',
            title='Genomsnittlig Beläggningsgrad per Timme (Vardag vs Helg)',
            labels={'Hour': 'Timme på Dygnet', 'Utilization': 'Genomsnittlig Beläggning', 'Day_Type': 'Dagtyp'},
            markers=True, color_discrete_map={'Vardag': '#27ae60', 'Helg': '#f39c12'}  # Changed to green for weekdays
        )
        fig_avg_hourly_util.update_layout(yaxis_tickformat=".0%", xaxis_dtick=1)
        figures['avg_hourly_utilization_line'] = fig_avg_hourly_util
    else:
        figures['avg_hourly_utilization_line'] = go.Figure().update_layout(title_text='Genomsnittlig Beläggningsgrad per Timme (Ingen data)')

    # 4c. Aggregated Weekday Utilization (Bar chart) - Changed to green
    if not metrics['avg_weekday_utilization'].empty:
        fig_avg_weekday_util = px.bar(
            metrics['avg_weekday_utilization'], x='Weekday', y='Utilization',
            title='Genomsnittlig Beläggningsgrad per Veckodag',
            labels={'Weekday': 'Veckodag', 'Utilization': 'Genomsnittlig Beläggning'},
            color='Utilization', color_continuous_scale=px.colors.sequential.Greens  # Changed to green
        )
        fig_avg_weekday_util.update_layout(yaxis_tickformat=".0%")
        figures['avg_weekday_utilization_bar'] = fig_avg_weekday_util
    else:
        figures['avg_weekday_utilization_bar'] = go.Figure().update_layout(title_text='Genomsnittlig Beläggningsgrad per Veckodag (Ingen data)')
        
    # 5. Heatmap: Monthly Utilization per area - Changed to green
    if not metrics['utilization'].empty and 'Utilization' in metrics['utilization'].columns:
        pivot_util = metrics['utilization'].pivot_table(
            index='Område', columns='Year_Month', values='Utilization'
        ).fillna(0)
        
        # Find min/max for consistent gradient
        min_val = pivot_util.values.min()
        max_val = pivot_util.values.max()
        
        fig_util = px.imshow(
            pivot_util, labels=dict(x='Månad', y='Område', color='Beläggning (%)'),
            x=pivot_util.columns, y=pivot_util.index, 
            color_continuous_scale=px.colors.sequential.Greens,  # Changed to green
            title='Månatlig Beläggningsgrad per Område (Baserat på Aktiva Uttagstimmar)',
            zmin=min_val, zmax=max(0.01, max_val)  # Set fixed min/max for consistent gradient
        )
        fig_util.update_layout(xaxis_tickangle=-45, coloraxis_colorbar=dict(tickformat='.0%'))
        figures['monthly_utilization_area_heatmap'] = fig_util
    else:
        figures['monthly_utilization_area_heatmap'] = go.Figure().update_layout(title_text='Månatlig Beläggningsgrad per Område (Ingen data)')
        
    # 6. Box plot: kWh per OUTLET per month per area (Distribution of individual outlet performance)
    if not metrics['kwh_outlet_month_area'].empty:
        fig_kwh_outlet_dist = px.box(
            metrics['kwh_outlet_month_area'], x='Område', y='Laddat (kWh)', color='Year_Month',
            title='Distribution av Energi (kWh) per Uttag (Månadsvis per Område)',
            labels={'Laddat (kWh)': 'Energi per Uttag (kWh)', 'Område': 'Område', 'Year_Month': 'Månad'},
            color_discrete_sequence=px.colors.qualitative.Vivid
        )
        fig_kwh_outlet_dist.update_layout(xaxis_tickangle=-45)
        figures['kwh_per_outlet_distribution'] = fig_kwh_outlet_dist
    else:
        figures['kwh_per_outlet_distribution'] = go.Figure().update_layout(title_text='Distribution av Energi (kWh) per Uttag (Ingen data)')

    # 7. Histogram: Session Duration - CHANGED TO HOURS AND REMOVE OUTLIERS OVER 12 HOURS
    if not sessions_df.empty and 'Duration_Hours' in sessions_df.columns:
        # Filter outliers - sessions longer than 12 hours
        filtered_sessions = sessions_df[sessions_df['Duration_Hours'] <= 12]
        
        fig_session_duration_hist = px.histogram(
            filtered_sessions, x='Duration_Hours', nbins=50,
            title='Distribution av Sessionslängd (Timmar)',
            labels={'Duration_Hours': 'Sessionslängd (Timmar)', 'count': 'Antal Sessioner'},
            marginal="box",  # adds a box plot above histogram
            color_discrete_sequence=['#27ae60']  # Use green color
        )
        fig_session_duration_hist.update_layout(yaxis_title="Antal Sessioner")
        figures['session_duration_histogram'] = fig_session_duration_hist
    else:
        figures['session_duration_histogram'] = go.Figure().update_layout(title_text='Distribution av Sessionslängd (Ingen data)')

    # 8. Box Plot: Energy per Session by Area
    if not sessions_df.empty and 'Laddat (kWh)' in sessions_df.columns:
        fig_kwh_per_session_area_box = px.box(
            sessions_df, x='Område', y='Laddat (kWh)', color='Område',
            title='Energi (kWh) per Session fördelat på Område',
            labels={'Laddat (kWh)': 'Energi per Session (kWh)', 'Område': 'Område'},
            color_discrete_sequence=px.colors.qualitative.Safe
        )
        fig_kwh_per_session_area_box.update_layout(xaxis_tickangle=-45)
        figures['kwh_per_session_area_box'] = fig_kwh_per_session_area_box
    else:
        figures['kwh_per_session_area_box'] = go.Figure().update_layout(title_text='Energi (kWh) per Session (Ingen data)')

    return figures

# Function to generate HTML report
def generate_html_report(metrics, figures, selected_graph_keys, selected_areas, date_range_text):
    """
    Generate HTML report with selected graphs and metrics
    """
    # Create header information
    header_info_list = []
    for area in selected_areas:
        area_outlets = metrics['outlets_per_area'][metrics['outlets_per_area']['Område'] == area]['Number_of_Outlets'].values
        if len(area_outlets) > 0:
            header_info_list.append({
                'Anläggningsnamn': area,
                'Antal_Uttag': area_outlets[0]
            })
    
    current_date_str_report = datetime.now().strftime('%Y-%m-%d %H:%M')
    
    # Create HTML header
    html_header_parts = ["<div class='report-title'>ChargeNode - Analysrapport</div>"]
    html_header_parts.append(f"<div class='report-subtitle'>{date_range_text}</div>")
    
    if header_info_list:
        html_header_parts.append("<ul class='facility-list'>")
        for facility_info in header_info_list:
            html_header_parts.append(f"<li>{facility_info['Anläggningsnamn']} ({facility_info['Antal_Uttag']} uttag)</li>")
        html_header_parts.append("</ul>")
    else:
        if selected_areas:
            html_header_parts.append(f"<p class='facility-list-empty'>Områden valda: {', '.join(selected_areas)} (ingen sessionsdata för dessa i perioden, eller inga uttag registrerade).</p>")
        else:
            html_header_parts.append("<p class='facility-list-empty'>Inga områden valda för rapporten.</p>")
    
    html_header_parts.append(f"<p class='generation-date'>Genererad: {current_date_str_report}</p>")

    # Build HTML content
    html_content_parts = [f"""
    <!DOCTYPE html>
    <html lang="sv">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>ChargeNode - Laddningsrapport</title>
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
            
            body {{
                font-family: 'Inter', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 0;
                padding: 0;
                background-color: #f4f7f6;
                color: #333;
                line-height: 1.6;
            }}
            
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                padding: 2rem;
                background-color: #fff;
                box-shadow: 0 0 20px rgba(0,0,0,0.1);
                border-radius: 10px;
                overflow: auto;
            }}
            
            .header-logo {{
                position: absolute;
                top: 2rem;
                right: 2rem;
                width: 180px;
                height: auto;
            }}
            
            .report-header {{
                position: relative;
                padding-bottom: 1.5rem;
                margin-bottom: 2rem;
                border-bottom: 2px solid #27ae60;
                text-align: left;
            }}
            
            .report-title {{
                font-size: 2.2rem;
                font-weight: 700;
                color: #27ae60;
                margin-bottom: 0.5rem;
                line-height: 1.2;
            }}
            
            .report-subtitle {{
                font-size: 1.4rem;
                color: #555;
                margin-bottom: 1rem;
                font-weight: 400;
            }}
            
            .facility-list {{
                list-style-type: none;
                padding-left: 0;
                margin-bottom: 1rem;
                font-size: 1.1rem;
            }}
            
            .facility-list li {{
                margin-bottom: 0.3rem;
                display: inline-block;
                margin-right: 1.5rem;
                background-color: #f1f9f1;
                padding: 0.4rem 0.8rem;
                border-radius: 4px;
                border-left: 3px solid #27ae60;
            }}
            
            .facility-list-empty {{
                font-style: italic;
                font-size: 1rem;
                color: #777;
            }}
            
            .generation-date {{
                font-size: 0.9rem;
                color: #777;
            }}
            
            .metrics-container {{
                display: flex;
                flex-wrap: wrap;
                justify-content: space-between;
                margin-bottom: 2rem;
                gap: 20px;
            }}
            
            .metric-card {{
                flex: 1 1 calc(25% - 20px);
                min-width: 200px;
                background-color: #f8f9fa;
                border-radius: 8px;
                padding: 1.5rem;
                box-shadow: 0 2px 8px rgba(0,0,0,0.08);
                text-align: center;
                height: 150px;
                display: flex;
                flex-direction: column;
                justify-content: center;
                border-top: 4px solid #27ae60;
                transition: transform 0.2s, box-shadow 0.2s;
            }}
            
            .metric-card:hover {{
                transform: translateY(-5px);
                box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            }}
            
            .metric-value {{
                font-size: 2.5rem;
                font-weight: 700;
                color: #27ae60;
                margin-bottom: 0.5rem;
                line-height: 1;
            }}
            
            .metric-label {{
                font-size: 1rem;
                color: #555;
                font-weight: 500;
            }}
            
            .graph-section {{
                margin-bottom: 3rem;
                background-color: #fff;
                border-radius: 10px;
                padding: 1.5rem;
                box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            }}
            
            .graph-title {{
                font-size: 1.5rem;
                color: #27ae60;
                margin-bottom: 1rem;
                padding-bottom: 0.5rem;
                border-bottom: 1px solid #e0e0e0;
                font-weight: 600;
            }}
            
            .graph-description {{
                font-size: 1rem;
                color: #666;
                margin-bottom: 1rem;
                font-style: normal;
                line-height: 1.6;
                max-width: 80%;
            }}
            
            .plotly-graph-div {{
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 1rem;
                background-color: #fff;
                margin-bottom: 1rem;
                height: 600px;
                width: 100%;
                box-sizing: border-box;
            }}
            
            .print-button-container {{
                position: sticky;
                bottom: 2rem;
                right: 2rem;
                text-align: right;
                z-index: 100;
                margin: 2rem 0;
            }}
            
            .print-button {{
                background-color: #27ae60;
                color: white;
                padding: 0.8rem 1.5rem;
                font-size: 1rem;
                border: none;
                border-radius: 5px;
                cursor: pointer;
                transition: background-color 0.3s ease;
                font-weight: 500;
                display: inline-flex;
                align-items: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            }}
            
            .print-button:hover {{
                background-color: #219653;
            }}
            
            .print-button svg {{
                margin-right: 8px;
            }}
            
            .print-tips {{
                margin: 2rem 0;
                padding: 1rem 1.5rem;
                background-color: #f1f9f1;
                border-left: 4px solid #27ae60;
                font-size: 0.9rem;
                border-radius: 0 4px 4px 0;
            }}
            
            .print-tips h3 {{
                margin-top: 0;
                color: #27ae60;
                font-size: 1.1rem;
            }}
            
            .print-tips ul {{
                margin-bottom: 0;
                padding-left: 1.2rem;
            }}
            
            .print-tips li {{
                margin-bottom: 0.5rem;
            }}
            
            .print-tips li:last-child {{
                margin-bottom: 0;
            }}
            
            .footer {{
                margin-top: 3rem;
                padding-top: 1.5rem;
                border-top: 1px solid #e0e0e0;
                color: #777;
                font-size: 0.9rem;
                text-align: center;
            }}
            
            /* Responsive adjustments */
            @media (max-width: 992px) {{
                .container {{
                    padding: 1.5rem;
                    max-width: 95%;
                }}
                
                .metric-card {{
                    flex: 1 1 calc(50% - 20px);
                }}
                
                .graph-description {{
                    max-width: 100%;
                }}
            }}
            
            @media (max-width: 768px) {{
                .container {{
                    padding: 1rem;
                }}
                
                .report-title {{
                    font-size: 1.8rem;
                }}
                
                .report-subtitle {{
                    font-size: 1.2rem;
                }}
                
                .metric-card {{
                    flex: 1 1 100%;
                }}
                
                .header-logo {{
                    position: static;
                    display: block;
                    margin: 0 auto 1rem;
                }}
                
                .report-header {{
                    text-align: center;
                }}
            }}
            
            /* Print styles */
            @media print {{
                @page {{
                    size: A4;
                    margin: 1cm;
                }}
                
                body {{
                    margin: 0;
                    padding: 0;
                    background-color: #fff;
                    font-size: 11pt;
                }}
                
                .container {{
                    max-width: 100%;
                    margin: 0;
                    padding: 0;
                    box-shadow: none;
                    border: none;
                }}
                
                .print-button-container,
                .print-tips {{
                    display: none !important;
                }}
                
                .graph-section {{
                    page-break-inside: avoid;
                    page-break-after: always;
                    break-inside: avoid;
                    margin-bottom: 20pt;
                    box-shadow: none;
                    border: 1px solid #eee;
                }}
                
                .graph-section:last-child {{
                    page-break-after: auto;
                }}
                
                .plotly-graph-div {{
                    height: 400pt !important;
                    width: 100% !important;
                    border: none;
                }}
                
                .modebar {{
                    display: none !important;
                }}
                
                .metric-card {{
                    page-break-inside: avoid;
                    break-inside: avoid;
                    box-shadow: none;
                    border: 1px solid #eee;
                }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="report-header">
                <img src="https://chargenode.eu/wp-content/themes/chargenode/dist/assets/img/logo.svg" alt="ChargeNode Logo" class="header-logo">
                {''.join(html_header_parts)}
            </div>
            
            <div class="print-tips">
                <h3>Hur du sparar denna rapport som PDF:</h3>
                <ul>
                    <li>Klicka på "Skriv ut / Spara som PDF" knappen längst ner på sidan.</li>
                    <li>Välj "Spara som PDF" som destination i utskriftsdialogen.</li>
                    <li>I utskriftsinställningarna, se till att du har markerat "Bakgrundsgrafik" och ställ in marginalerna till "Minimal" eller "Inga".</li>
                    <li>Förhandsgranska för att kontrollera att alla grafer visas korrekt och klicka sedan på "Spara".</li>
                </ul>
            </div>
            
            <!-- Key Metrics Section -->
            <div class="graph-section">
                <h2 class="graph-title">Nyckeltal</h2>
                <div class="metrics-container">
                    <div class="metric-card">
                        <div class="metric-value">{metrics['area_count']}</div>
                        <div class="metric-label">Antal Områden</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">{metrics['total_outlets']}</div>
                        <div class="metric-label">Totalt Antal Uttag</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">{metrics['total_sessions']:,}</div>
                        <div class="metric-label">Totala Sessioner</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">{metrics['total_kwh']:,.2f}</div>
                        <div class="metric-label">Total Energi (kWh)</div>
                    </div>
                </div>
            </div>
    """
    ]

    # Explanatory descriptions for each graph
    graph_descriptions = {
        'outlets_per_area': "Visar det totala antalet unika ladduttag som har använts inom den valda tidsperioden och för de valda områdena. Detta hjälper att få en översikt över fördelningen av laddinfrastruktur.",
        'kwh_per_month_area': "Stapeldiagram som visar den totala laddade energin (kWh) per månad, uppdelat per valt område. Användbart för att följa energitrender över tid och mellan olika platser.",
        'revenue_per_month_area': "Stapeldiagram som visar den totala intäkten (exklusive moms) per månad, uppdelat per valt område. Ger insikt i ekonomiska aspekter över tid och per plats.",
        'avg_kwh_per_outlet': "Linjediagram som visar den genomsnittliga mängden energi (kWh) som varje aktivt uttag har levererat per månad, sett över alla valda områden. En ökande trend kan indikera effektivare användning eller längre sessioner.",
        'hourly_utilization_heatmap': "Heatmap som visualiserar beläggningsgraden för varje timme och dag. Mörkare grön färg indikerar högre beläggning. Detta ger en tydlig bild av användningsmönster över tid och hjälper till att identifiera perioder med hög efterfrågan.",
        'avg_hourly_utilization_line': "Linjediagram som visar den genomsnittliga beläggningsgraden fördelat per timme på dygnet, uppdelat på vardagar och helger. Detta hjälper att identifiera tidpunkter för mest aktiv användning.",
        'avg_weekday_utilization_bar': "Stapeldiagram som jämför den genomsnittliga beläggningsgraden för varje veckodag. Perfekt för att planera underhåll eller identifiera mönster i veckoanvändning.",
        'monthly_utilization_area_heatmap': "Heatmap som visar den beräknade månatliga beläggningsgraden per område. Detta ger insikt i hur effektivt laddinfrastrukturen används över tid i olika områden.",
        'session_duration_histogram': "Histogram som visar fördelningen av längden på laddningssessionerna (i timmar). Outliers över 12 timmar har filtrerats bort för att ge en tydligare bild av vanliga laddningsmönster.",
        'kwh_per_session_area_box': "Boxplot som illustrerar spridningen av laddad energi (kWh) per session, för varje valt område. Visar median, kvartiler och ger insikt i typiska laddningsvolymer.",
        'kwh_per_outlet_distribution': "Boxplot som visar fördelningen av total energi (kWh) som varje enskilt uttag har levererat under en månad, per område. Hjälper till att identifiera hög- och lågpresterande uttag."
    }

    # Add selected graphs
    plotly_js_added = False
    
    for key in selected_graph_keys:
        if key in figures:
            fig = figures[key]
            
            # Get graph display name and description
            graph_display_names = {
                'outlets_per_area': 'Antal Uttag per Område',
                'kwh_per_month_area': 'Energi per Månad och Område',
                'revenue_per_month_area': 'Omsättning per Månad och Område',
                'avg_kwh_per_outlet': 'Genomsnittlig kWh per Uttag per Månad',
                'hourly_utilization_heatmap': 'Timvis Beläggningsgrad (Heatmap)',
                'avg_hourly_utilization_line': 'Genomsnittlig Beläggning per Timme',
                'avg_weekday_utilization_bar': 'Genomsnittlig Beläggning per Veckodag',
                'monthly_utilization_area_heatmap': 'Månatlig Beläggning per Område',
                'session_duration_histogram': 'Distribution av Sessionslängd',
                'kwh_per_session_area_box': 'Energi per Session per Område',
                'kwh_per_outlet_distribution': 'Distribution av Energi per Uttag'
            }
            
            graph_title = fig.layout.title.text if hasattr(fig.layout, 'title') and fig.layout.title else graph_display_names.get(key, key)
            graph_description = graph_descriptions.get(key, "")
            
            # Optimize graph for better display
            fig.update_layout(
                autosize=True,
                height=600,
                margin=dict(l=80, r=40, t=100, b=80),
                font=dict(size=14),
                title=dict(
                    font=dict(size=18),
                    y=0.95
                )
            )
            
            html_content_parts.append(f"<div class='graph-section'><h2 class='graph-title'>{graph_title}</h2>")
            if graph_description:
                html_content_parts.append(f"<p class='graph-description'>{graph_description}</p>")
            
            try:
                # Include Plotly.js only once
                if not plotly_js_added:
                    graph_html = fig.to_html(full_html=False, include_plotlyjs='cdn', config={'displayModeBar': False})
                    plotly_js_added = True
                else:
                    graph_html = fig.to_html(full_html=False, include_plotlyjs=False, config={'displayModeBar': False})
                html_content_parts.append(graph_html)
            except Exception as e:
                html_content_parts.append(f"<p><strong>Fel:</strong> Kunde inte rendera grafen '{graph_title}'. ({e})</p>")
            
            html_content_parts.append("</div>")
    
    # Add print button and footer
    html_content_parts.append("""
        <div class="print-button-container">
            <button class="print-button" onclick="window.print()">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <polyline points="6 9 6 2 18 2 18 9"></polyline>
                    <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path>
                    <rect x="6" y="14" width="12" height="8"></rect>
                </svg>
                Skriv ut / Spara som PDF
            </button>
        </div>
        
        <div class="footer">
            <p>ChargeNode Rapport Generator © 2025. Alla data är konfidentiella och endast för internt bruk.</p>
        </div>
    """)
    
    # Close HTML structure
    html_content_parts.append("</div></body></html>")
    
    # Join HTML parts
    final_html = "".join(html_content_parts)
    return final_html

# --- Main application logic ---
if sessions_file is not None: # Overview file is now optional
    try:
        st.info("Läser in och förbereder data...")
        sessions_df_original = pd.read_excel(sessions_file)
        overview_df_original = pd.read_excel(overview_file) if overview_file else None
        
        st.sidebar.success("Data inläst!")
        
        st.sidebar.header("Filter")
        
        sessions_df_processed, overview_df_processed = preprocess_data(sessions_df_original.copy(), overview_df_original.copy() if overview_df_original is not None else None) # Use copies
        
        if sessions_df_processed.empty:
            st.error("Ingen sessionsdata hittades eller all data filtrerades bort under förbehandling. Kontrollera din fil.")
            st.stop()

        all_areas = sorted(sessions_df_processed['Område'].unique())
        selected_areas = st.sidebar.multiselect("Välj Områden", options=all_areas, default=all_areas)
        
        # Date range filter
        min_data_date = sessions_df_processed['Startad'].min().date()
        max_data_date = sessions_df_processed['Startad'].max().date()

        selected_date_range = st.sidebar.date_input(
            "Välj Datumintervall",
            value=(min_data_date, max_data_date),
            min_value=min_data_date,
            max_value=max_data_date,
            key="date_range_filter"
        )

        # Generate date range text for reports
        if len(selected_date_range) == 2:
            start_date_filter = pd.to_datetime(selected_date_range[0])
            end_date_filter = pd.to_datetime(selected_date_range[1]) + pd.Timedelta(days=1) # Include the whole end day
            filtered_sessions = sessions_df_processed[
                (sessions_df_processed['Område'].isin(selected_areas)) &
                (sessions_df_processed['Startad'] >= start_date_filter) &
                (sessions_df_processed['Startad'] < end_date_filter)
            ]
            date_range_text = f"Period: {selected_date_range[0].strftime('%Y-%m-%d')} till {selected_date_range[1].strftime('%Y-%m-%d')}"
        else: # Default to all dates if range is not correctly selected
            filtered_sessions = sessions_df_processed[sessions_df_processed['Område'].isin(selected_areas)]
            date_range_text = f"Period: {min_data_date.strftime('%Y-%m-%d')} till {max_data_date.strftime('%Y-%m-%d')}"
        
        if filtered_sessions.empty:
            st.warning("Ingen data matchar de valda filtren. Prova att ändra filterinställningarna.")
        else:
            st.info("Beräknar nyckeltal...")
            metrics = calculate_metrics(filtered_sessions, overview_df_processed)
            
            st.info("Skapar visualiseringar...")
            figures = create_visualizations(metrics, filtered_sessions) # Pass filtered_sessions for detailed plots
            
            # Clear info messages
            st.empty() # Clears the last message, may need more if they stack
            st.empty()
            
            # --- Display Dashboard ---
            tab_titles = [
                "📊 Nyckeltal & Översikt", 
                "⏱️ Beläggningsanalys", 
                "💡 Energiförbrukning", 
                "🔌 Sessionsanalys"
            ]
            tab1, tab2, tab3, tab4 = st.tabs(tab_titles)
            
            with tab1:
                st.header("Nyckeltal")
                cols = st.columns(5)  # Uppdatera till 5 kolumner

                key_metrics_display = {
                    "Antal Områden": metrics.get('area_count', 'N/A'),
                    "Totalt Antal Uttag": f"{metrics.get('total_outlets', 'N/A'):,}",
                    "Totala Sessioner": f"{metrics.get('total_sessions', 'N/A'):,}",
                    "Total Energi (kWh)": f"{metrics.get('total_kwh', 0):,.2f}",
                    "Total Omsättning (SEK)": f"{metrics.get('total_revenue', 0):,.2f}"
                }

                for i, (label, value) in enumerate(key_metrics_display.items()):
                    with cols[i % 4]:
                         st.markdown(f"""
                            <div class="metric-card">
                                <div class="metric-value">{value}</div>
                                <div class="metric-label">{label}</div>
                            </div>""", unsafe_allow_html=True)
                st.markdown("---")
                st.subheader(figures['outlets_per_area'].layout.title.text)
                st.caption("Visar det totala antalet unika ladduttag som har använts inom den valda tidsperioden och för de valda områdena.")
                st.plotly_chart(figures['outlets_per_area'], use_container_width=True)

            with tab2:
                st.header("Beläggningsanalys")
                st.subheader(figures['hourly_utilization_heatmap'].layout.title.text)
                st.caption("Heatmap som visualiserar beläggningsgraden (andelen aktiva uttag av totalt antal uttag i de valda områdena) för varje timme och dag. Mörkare grön färg indikerar högre beläggning. Detta hjälper till att identifiera mönster och tider med hög/låg efterfrågan.")
                st.plotly_chart(figures['hourly_utilization_heatmap'], use_container_width=True)
                
                st.subheader(figures['avg_hourly_utilization_line'].layout.title.text)
                st.caption("Linjediagram som visar den genomsnittliga beläggningsgraden fördelat per timme på dygnet, uppdelat på vardagar och helger. Toppar indikerar perioder med högst genomsnittlig användning.")
                st.plotly_chart(figures['avg_hourly_utilization_line'], use_container_width=True)

                st.subheader(figures['avg_weekday_utilization_bar'].layout.title.text)
                st.caption("Stapeldiagram som jämför den genomsnittliga beläggningsgraden för varje veckodag. Ger en snabb överblick över vilka dagar som har högst respektive lägst användning.")
                st.plotly_chart(figures['avg_weekday_utilization_bar'], use_container_width=True)

                st.subheader(figures['monthly_utilization_area_heatmap'].layout.title.text)
                st.caption("Heatmap som visar den beräknade månatliga beläggningsgraden per område. Beläggningen är här definierad som totala antalet timmar uttagen varit aktiva i förhållande till det totala antalet möjliga drifttimmar för alla uttag i området under månaden.")
                st.plotly_chart(figures['monthly_utilization_area_heatmap'], use_container_width=True)

            with tab3:
                st.header("Energiförbrukning och Omsättning")
                st.subheader(figures['kwh_per_month_area'].layout.title.text)
                st.caption("Stapeldiagram som visar den totala laddade energin (kWh) per månad, uppdelat per valt område. Användbart för att följa energitrender över tid och mellan olika platser.")
                st.plotly_chart(figures['kwh_per_month_area'], use_container_width=True)

                st.subheader(figures['revenue_per_month_area'].layout.title.text)
                st.caption("Stapeldiagram som visar den totala Omsättningen (exklusive moms) per månad, uppdelat per valt område. Ger insikt i ekonomiska aspekter över tid och per plats.")
                st.plotly_chart(figures['revenue_per_month_area'], use_container_width=True)

                st.subheader(figures['avg_kwh_per_outlet'].layout.title.text)
                st.caption("Linjediagram som visar den genomsnittliga mängden energi (kWh) som varje aktivt uttag har levererat per månad, sett över alla valda områden. En ökande trend kan indikera effektivare användning eller längre sessioner per uttag.")
                st.plotly_chart(figures['avg_kwh_per_outlet'], use_container_width=True)
            
            with tab4:
                st.header("Sessionsanalys")
                st.subheader(figures['session_duration_histogram'].layout.title.text)
                st.caption("Histogram som visar fördelningen av längden på laddningssessionerna (i timmar). Outliers över 12 timmar har filtrerats bort. Detta ger en bild av hur länge användare typiskt laddar sina fordon.")
                st.plotly_chart(figures['session_duration_histogram'], use_container_width=True)

                st.subheader(figures['kwh_per_session_area_box'].layout.title.text)
                st.caption("Boxplot som illustrerar spridningen (median, kvartiler, och extremvärden) av laddad energi (kWh) per session, för varje valt område. Hjälper till att förstå typisk energimängd per laddning och variationen mellan områden.")
                st.plotly_chart(figures['kwh_per_session_area_box'], use_container_width=True)

                st.subheader(figures['kwh_per_outlet_distribution'].layout.title.text)
                st.caption("Boxplot som visar fördelningen av total energi (kWh) som varje enskilt uttag har levererat under en månad, per område. Detta kan hjälpa till att identifiera uttag som är särskilt hög- eller lågpresterande inom ett område.")
                st.plotly_chart(figures['kwh_per_outlet_distribution'], use_container_width=True)

            # --- Export Section ---
            st.sidebar.markdown("---")
            st.sidebar.header("Exportera Rapport")

            # Initialize session state for graph selection
            if 'default_selected_graphs' not in st.session_state:
                st.session_state.default_selected_graphs = list(figures.keys())
            if 'all_graph_keys' not in st.session_state:
                st.session_state.all_graph_keys = list(figures.keys())
            if 'select_all_toggle' not in st.session_state:
                st.session_state.select_all_toggle = True

            # Section for HTML report
            st.sidebar.subheader("HTML Rapport")
            
            # Buttons for selecting/deselecting all graphs
            col1, col2 = st.sidebar.columns(2)
            if col1.button("Markera alla", key="select_all_html"):
                st.session_state.select_all_toggle = True
                st.session_state.default_selected_graphs = st.session_state.all_graph_keys
                st.rerun()

            if col2.button("Avmarkera alla", key="deselect_all_html"):
                st.session_state.select_all_toggle = False
                st.session_state.default_selected_graphs = []
                st.rerun()
            
            # If select_all_toggle is true and no graphs are selected, select all
            if st.session_state.select_all_toggle and not st.session_state.default_selected_graphs and st.session_state.all_graph_keys:
                st.session_state.default_selected_graphs = st.session_state.all_graph_keys

            # Map of more readable graph names for display
            graph_display_names = {
                'outlets_per_area': 'Antal Uttag per Område',
                'kwh_per_month_area': 'Energi per Månad och Område',
                'revenue_per_month_area': 'Omsättning per Månad och Område',
                'avg_kwh_per_outlet': 'Genomsnittlig kWh per Uttag per Månad',
                'hourly_utilization_heatmap': 'Timvis Beläggningsgrad (Heatmap)',
                'avg_hourly_utilization_line': 'Genomsnittlig Beläggning per Timme',
                'avg_weekday_utilization_bar': 'Genomsnittlig Beläggning per Veckodag',
                'monthly_utilization_area_heatmap': 'Månatlig Beläggning per Område',
                'session_duration_histogram': 'Distribution av Sessionslängd',
                'kwh_per_session_area_box': 'Energi per Session per Område',
                'kwh_per_outlet_distribution': 'Distribution av Energi per Uttag'
            }

            # Multiselect for graph selection
            selected_graph_keys = st.sidebar.multiselect(
                "Välj grafer för HTML-rapport:",
                options=st.session_state.all_graph_keys,
                format_func=lambda key: graph_display_names.get(key, str(key)),
                default=st.session_state.default_selected_graphs,
                key="graph_selector_html"
            )
            
            # Update default_selected_graphs to remember selection
            st.session_state.default_selected_graphs = selected_graph_keys

            # Generate HTML button
            if st.sidebar.button("Generera HTML-rapport"):
                if not selected_graph_keys:
                    st.sidebar.warning("Vänligen välj minst en graf att inkludera i rapporten.")
                else:
                    # Generate HTML report
                    final_html = generate_html_report(metrics, figures, selected_graph_keys, selected_areas, date_range_text)
                    
                    # Create download button
                    report_filename = f"ChargeNode_Rapport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
                    st.sidebar.download_button(
                        label="📥 Ladda ner HTML-rapporten",
                        data=final_html,
                        file_name=report_filename,
                        mime="text/html",
                        key="download_html_button"
                    )
                    st.sidebar.success(f"HTML-rapport '{report_filename}' genererad och redo för nedladdning!")
                    
                    # Show preview (optional)
                    with st.sidebar.expander("Förhandsgranska HTML-rapport"):
                        st.markdown(f"<iframe srcdoc='{final_html}' width='100%' height='400px'></iframe>", unsafe_allow_html=True)
            
            # PDF instructions
            with st.sidebar.expander("Så här skapar du en PDF-rapport"):
                st.markdown("""
                ### Skapa PDF från HTML-rapporten
                
                1. Generera och ladda ner HTML-rapporten först genom att klicka på knappen "Generera HTML-rapport"
                2. Öppna HTML-filen i Chrome, Firefox eller Edge
                3. Tryck på "Skriv ut" (Ctrl+P eller ⌘+P) eller använd knappen längst ner i rapporten
                4. Välj "Spara som PDF" som destination/skrivare
                5. Kontrollera att "Bakgrundsgrafik" är markerat i utskriftsalternativen
                6. Ställ in marginalerna till "Minimal" eller "Inga" för bästa resultat
                7. Klicka på "Spara" för att skapa PDF-filen
                
                Detta ger dig en professionell rapport med ChargeNode grafik och alla valda visualiseringar.
                """)
    
    except FileNotFoundError:
        st.error(f"Fel: En eller båda Excel-filerna kunde inte hittas. Kontrollera filnamn och sökvägar.")
    except ValueError as ve:
        st.error(f"Värdefel vid databehandling: {ve}. Kontrollera att dina datafiler har förväntat format och innehåll, särskilt datum och numeriska värden.")
    except Exception as e:
        st.error(f"Ett oväntat fel inträffade: {e}")
        st.exception(e) # Shows full traceback for debugging

elif sessions_file is None and overview_file is not None:
    st.warning("Ladda upp 'Sessions.xlsx' för att påbörja analysen. 'Overview.xlsx' är valfri.")
else:
    st.info("Vänligen ladda upp 'Sessions.xlsx' för att visualisera data. 'Overview.xlsx' är valfri men kan ge ytterligare kontext.")
