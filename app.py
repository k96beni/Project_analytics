import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import tempfile
import os
from datetime import datetime
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
import io
import base64
from PIL import Image as PILImage # Alias för att undvika namnkonflikt
import calendar

# Set page configuration
st.set_page_config(
    page_title="Charging Outlets Dashboard",
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
        color: #2c3e50; /* Adjust color as needed */
    }
    /* Add more custom styles if needed */
</style>
""", unsafe_allow_html=True)

# --- Data Loading and Processing ---
def load_data(uploaded_files):
    all_data = []
    for uploaded_file in uploaded_files:
        try:
            # Försök läsa som Excel först
            df = pd.read_excel(uploaded_file, sheet_name="Laddsessioner")
        except Exception as e_excel:
            # Om Excel-läsning misslyckas, försök läsa som CSV
            try:
                # Återställ filpekaren för CSV-läsning
                uploaded_file.seek(0)
                # Lägger till encoding='utf-8-sig' för att hantera eventuell BOM (Byte Order Mark)
                # och specificerar quotechar för att hantera eventuella citattecken i datan
                df = pd.read_csv(uploaded_file, sep=';', decimal=',', encoding='utf-8-sig', quotechar='"', quoting=0) # quoting=0 (QUOTE_MINIMAL)
            except Exception as e_csv:
                st.error(f"Fel vid läsning av fil {uploaded_file.name}: Varken Excel- eller CSV-format fungerade. Detaljer: Excel ({e_excel}), CSV ({e_csv})")
                continue
        all_data.append(df)
    
    if not all_data:
        return pd.DataFrame() # Returnera tom DataFrame om ingen data kunde laddas
        
    combined_df = pd.concat(all_data, ignore_index=True)

    # Datatvätt och transformation
    # Konvertera datumkolumner
    if 'Starttid' in combined_df.columns:
        # Använd errors='coerce' för att omvandla ogiltiga datumformat till NaT (Not a Time)
        combined_df['Starttid'] = pd.to_datetime(combined_df['Starttid'], errors='coerce')
        if pd.api.types.is_datetime64_any_dtype(combined_df['Starttid']): # Kontrollera om kolumnen är datetime
            # Säkerställ att alla datum är UTC-medvetna för konsistens
            if combined_df['Starttid'].dt.tz is None: # Om serien är naiv (ingen tidszonsinfo)
                if not combined_df['Starttid'].isnull().all(): # Om det finns några icke-NaT värden
                    combined_df['Starttid'] = combined_df['Starttid'].dt.tz_localize('UTC', ambiguous='infer')
            else: # Om serien redan är tidszonsmedveten
                combined_df['Starttid'] = combined_df['Starttid'].dt.tz_convert('UTC')

    if 'Sluttid' in combined_df.columns:
        combined_df['Sluttid'] = pd.to_datetime(combined_df['Sluttid'], errors='coerce')
        if pd.api.types.is_datetime64_any_dtype(combined_df['Sluttid']): # Kontrollera om kolumnen är datetime
            if combined_df['Sluttid'].dt.tz is None: # Om serien är naiv
                if not combined_df['Sluttid'].isnull().all():
                    combined_df['Sluttid'] = combined_df['Sluttid'].dt.tz_localize('UTC', ambiguous='infer')
            else: # Om serien redan är tidszonsmedveten
                combined_df['Sluttid'] = combined_df['Sluttid'].dt.tz_convert('UTC')

    # Varna för NaT-värden efter konverteringsförsök (innan rader tas bort)
    nat_start_count_initial = 0
    if 'Starttid' in combined_df.columns:
        nat_start_count_initial = combined_df['Starttid'].isnull().sum()
        if nat_start_count_initial > 0:
            st.warning(f"{nat_start_count_initial} värden i 'Starttid' kunde inte tolkas som giltiga datum och har markerats som NaT (Not a Time).")
    
    nat_slut_count_initial = 0
    if 'Sluttid' in combined_df.columns:
        nat_slut_count_initial = combined_df['Sluttid'].isnull().sum()
        if nat_slut_count_initial > 0:
            st.warning(f"{nat_slut_count_initial} värden i 'Sluttid' kunde inte tolkas som giltiga datum och har markerats som NaT (Not a Time).")

    # Ta bort rader där Starttid eller Sluttid är NaT, eftersom de är kritiska
    initial_row_count = len(combined_df)
    combined_df.dropna(subset=['Starttid', 'Sluttid'], inplace=True)
    rows_dropped = initial_row_count - len(combined_df)
    if rows_dropped > 0:
        st.info(f"{rows_dropped} rader togs bort på grund av ogiltiga eller saknade värden i 'Starttid' eller 'Sluttid'.")

    # Konvertera numeriska kolumner, tvinga fel till NaN
    numeric_cols = ['Start Grund (SoC)', 'Slut Grund (SoC)', 'Start Meter (kWh)', 'Slut Meter (kWh)', 'Debiterad Energi (kWh)']
    for col in numeric_cols:
        if col in combined_df.columns:
            # Ersätt kommatecken med punkter om kolumnen är av objekttyp (sträng)
            if combined_df[col].dtype == 'object':
                 combined_df[col] = combined_df[col].str.replace(',', '.', regex=False)
            combined_df[col] = pd.to_numeric(combined_df[col], errors='coerce')

    # Beräkna varaktighet om Starttid och Sluttid är tillgängliga och giltiga
    if 'Starttid' in combined_df.columns and 'Sluttid' in combined_df.columns:
        # Säkerställ att båda kolumnerna är datetime och UTC-medvetna innan subtraktion
        if pd.api.types.is_datetime64_any_dtype(combined_df['Starttid']) and \
           pd.api.types.is_datetime64_any_dtype(combined_df['Sluttid']):
            combined_df['Varaktighet (minuter)'] = (combined_df['Sluttid'] - combined_df['Starttid']).dt.total_seconds() / 60
            combined_df['Varaktighet (minuter)'] = combined_df['Varaktighet (minuter)'].apply(lambda x: x if pd.notnull(x) and x >= 0 else 0) # Behåll 0-durationer
        else:
            combined_df['Varaktighet (minuter)'] = 0 # Eller np.nan om det är mer passande
            st.warning("Kunde inte beräkna 'Varaktighet (minuter)' då 'Starttid' eller 'Sluttid' inte är i korrekt datumformat efter rensning.")


    # Extrahera timme och veckodag etc.
    if 'Starttid' in combined_df.columns and pd.api.types.is_datetime64_any_dtype(combined_df['Starttid']):
        combined_df['Starttimme'] = combined_df['Starttid'].dt.hour
        combined_df['Startdag'] = combined_df['Starttid'].dt.day_name() # Engelsk dag
        # Konvertera engelska dagars namn till svenska
        day_map_en_sv = {
            "Monday": "Måndag", "Tuesday": "Tisdag", "Wednesday": "Onsdag",
            "Thursday": "Torsdag", "Friday": "Fredag", "Saturday": "Lördag", "Sunday": "Söndag"
        }
        combined_df['Startdag (SV)'] = combined_df['Startdag'].map(day_map_en_sv)

        combined_df['Månad'] = combined_df['Starttid'].dt.month 
        combined_df['År'] = combined_df['Starttid'].dt.year
    else:
        # Skapa tomma kolumner om Starttid inte är giltig, för att undvika fel senare
        for col_name in ['Starttimme', 'Startdag', 'Startdag (SV)', 'Månad', 'År']:
            if col_name not in combined_df.columns:
                combined_df[col_name] = np.nan if col_name in ['Starttimme', 'Månad', 'År'] else ""
        st.warning("Kunde inte extrahera tidsdetaljer (timme, dag, månad, år) då 'Starttid' saknas eller är i fel format.")
        
    return combined_df

# --- Plotting Functions ---
def plot_hourly_heatmap(df, area_filter):
    area_column = 'ChargePoint' 
    
    if area_filter == "All":
        filtered_df = df
        area_filter_disp = "alla områden"
    else:
        if area_column not in df.columns:
            st.warning(f"Kolumnen '{area_column}' som används för områdesfiltrering finns inte i datan.")
            return go.Figure()
        
        filtered_df = df[df[area_column] == area_filter]
        area_filter_disp = area_filter

    if filtered_df.empty:
        st.info(f"Ingen data tillgänglig för det valda området: {area_filter_disp} för att generera timvis värmekarta.")
        return go.Figure()

    # Använd 'Startdag (SV)' för svensk visning
    if 'Starttimme' in filtered_df.columns and 'Startdag (SV)' in filtered_df.columns and not filtered_df['Starttimme'].isnull().all() and not filtered_df['Startdag (SV)'].isnull().all():
        values_col = 'TransactionId' if 'TransactionId' in filtered_df.columns else filtered_df.columns[0]

        try:
            heatmap_data = pd.pivot_table(filtered_df, 
                                          values=values_col,
                                          index='Starttimme', 
                                          columns='Startdag (SV)', 
                                          aggfunc='count' if values_col == 'TransactionId' else 'size',
                                          fill_value=0)
        except Exception as e:
            st.error(f"Kunde inte skapa pivot-tabell för värmekarta: {e}")
            return go.Figure()

        if heatmap_data.empty:
            st.info(f"Ingen data att visa i värmekartan för {area_filter_disp} efter pivotering.")
            return go.Figure()

        # Ordning för svenska dagar
        days_order_sv = ["Måndag", "Tisdag", "Onsdag", "Torsdag", "Fredag", "Lördag", "Söndag"]
        heatmap_data = heatmap_data.reindex(columns=[day for day in days_order_sv if day in heatmap_data.columns])
        
        if heatmap_data.empty or heatmap_data.shape[1] == 0:
            st.info(f"Ingen data för relevanta veckodagar att visa i värmekartan för {area_filter_disp}.")
            return go.Figure()

        fig = px.imshow(heatmap_data,
                        labels=dict(x="Veckodag", y="Timme på dygnet", color="Antal Laddsessioner"),
                        x=heatmap_data.columns,
                        y=heatmap_data.index,
                        text_auto=True,
                        aspect="auto",
                        color_continuous_scale=px.colors.sequential.Viridis)
        
        fig.update_layout(
            title=f"Timvis Användning av Ladduttag ({area_filter_disp})",
            xaxis_title="Veckodag",
            yaxis_title="Timme på dygnet",
            height=700 
        )
        return fig
    else:
        missing_cols = []
        if 'Starttimme' not in filtered_df.columns or filtered_df['Starttimme'].isnull().all():
            missing_cols.append("'Starttimme'")
        if 'Startdag (SV)' not in filtered_df.columns or filtered_df['Startdag (SV)'].isnull().all():
            missing_cols.append("'Startdag (SV)'")
        st.warning(f"Nödvändiga och giltiga data i kolumnerna {', '.join(missing_cols)} saknas för att skapa värmekartan.")
        return go.Figure()

def plot_energy_consumption_trends(df, area_filter):
    area_column = 'ChargePoint'
    if area_filter == "All":
        filtered_df = df
        area_filter_disp = "alla områden"
    else:
        if area_column not in df.columns:
            st.warning(f"Kolumnen '{area_column}' som används för områdesfiltrering finns inte i datan.")
            return go.Figure()
        filtered_df = df[df[area_column] == area_filter]
        area_filter_disp = area_filter

    if filtered_df.empty:
        st.info(f"Ingen data tillgänglig för {area_filter_disp} för att visa energiförbrukningstrender.")
        return go.Figure()

    if 'Starttid' in filtered_df.columns and 'Debiterad Energi (kWh)' in filtered_df.columns and \
       pd.api.types.is_datetime64_any_dtype(filtered_df['Starttid']) and \
       not filtered_df['Debiterad Energi (kWh)'].isnull().all():
        
        # Säkerställ att 'Starttid' är sorterad för tidsserieplottning
        filtered_df = filtered_df.sort_values(by='Starttid')
        
        # Aggregera per dag för en tydligare trend
        daily_energy = filtered_df.set_index('Starttid').resample('D')['Debiterad Energi (kWh)'].sum().reset_index()
        
        if daily_energy.empty:
            st.info(f"Ingen aggregerad daglig energidata tillgänglig för {area_filter_disp}.")
            return go.Figure()

        fig = px.line(daily_energy, x='Starttid', y='Debiterad Energi (kWh)',
                      title=f"Trend för Energiförbrukning ({area_filter_disp})",
                      labels={'Starttid': 'Datum', 'Debiterad Energi (kWh)': 'Total Debiterad Energi (kWh)'})
        fig.update_layout(height=500)
        return fig
    else:
        st.warning("Nödvändiga kolumner ('Starttid', 'Debiterad Energi (kWh)') eller giltig data saknas för energitrender.")
        return go.Figure()

def plot_soc_distribution(df, area_filter):
    area_column = 'ChargePoint'
    if area_filter == "All":
        filtered_df = df
        area_filter_disp = "alla områden"
    else:
        if area_column not in df.columns:
            st.warning(f"Kolumnen '{area_column}' som används för områdesfiltrering finns inte i datan.")
            return go.Figure()
        filtered_df = df[df[area_column] == area_filter]
        area_filter_disp = area_filter
    
    if filtered_df.empty:
        st.info(f"Ingen data tillgänglig för {area_filter_disp} för att visa SoC-distribution.")
        return go.Figure()

    # Ta bort rader med NaN i SoC-kolumnerna för denna plott
    soc_df = filtered_df[['Start Grund (SoC)', 'Slut Grund (SoC)']].dropna()

    if soc_df.empty:
        st.info(f"Ingen giltig SoC-data tillgänglig för {area_filter_disp}.")
        return go.Figure()

    fig = go.Figure()
    if 'Start Grund (SoC)' in soc_df.columns and not soc_df['Start Grund (SoC)'].isnull().all():
        fig.add_trace(go.Histogram(x=soc_df['Start Grund (SoC)'], name='Start SoC (%)', nbinsx=20, marker_color='#1f77b4'))
    if 'Slut Grund (SoC)' in soc_df.columns and not soc_df['Slut Grund (SoC)'].isnull().all():
        fig.add_trace(go.Histogram(x=soc_df['Slut Grund (SoC)'], name='Slut SoC (%)', nbinsx=20, marker_color='#ff7f0e'))
    
    fig.update_layout(
        title_text=f'Distribution av Start- och Slut-SoC ({area_filter_disp})',
        xaxis_title_text='State of Charge (SoC %)',
        yaxis_title_text='Antal Sessioner',
        barmode='overlay',
        height=500
    )
    fig.update_traces(opacity=0.75)
    return fig

def plot_charging_duration_distribution(df, area_filter):
    area_column = 'ChargePoint'
    if area_filter == "All":
        filtered_df = df
        area_filter_disp = "alla områden"
    else:
        if area_column not in df.columns:
            st.warning(f"Kolumnen '{area_column}' som används för områdesfiltrering finns inte i datan.")
            return go.Figure()
        filtered_df = df[df[area_column] == area_filter]
        area_filter_disp = area_filter

    if filtered_df.empty or 'Varaktighet (minuter)' not in filtered_df.columns or filtered_df['Varaktighet (minuter)'].isnull().all():
        st.info(f"Ingen data för laddningsduration tillgänglig för {area_filter_disp}.")
        return go.Figure()
    
    # Filtrera bort orimligt långa sessioner om nödvändigt, t.ex. över 24 timmar = 1440 minuter
    duration_data = filtered_df[filtered_df['Varaktighet (minuter)'] < 1440]['Varaktighet (minuter)'].dropna()

    if duration_data.empty:
        st.info(f"Ingen giltig data för laddningsduration (under 24h) tillgänglig för {area_filter_disp}.")
        return go.Figure()

    fig = px.histogram(duration_data, nbins=30, title=f"Distribution av Laddningsduration ({area_filter_disp})")
    fig.update_layout(
        xaxis_title="Laddningsduration (minuter)",
        yaxis_title="Antal Sessioner",
        height=500
    )
    return fig
    
def plot_energy_vs_duration(df, area_filter):
    area_column = 'ChargePoint'
    if area_filter == "All":
        filtered_df = df
        area_filter_disp = "alla områden"
    else:
        if area_column not in df.columns:
            st.warning(f"Kolumnen '{area_column}' som används för områdesfiltrering finns inte i datan.")
            return go.Figure()
        filtered_df = df[df[area_column] == area_filter]
        area_filter_disp = area_filter

    if filtered_df.empty or \
       'Varaktighet (minuter)' not in filtered_df.columns or \
       'Debiterad Energi (kWh)' not in filtered_df.columns or \
       filtered_df['Varaktighet (minuter)'].isnull().all() or \
       filtered_df['Debiterad Energi (kWh)'].isnull().all():
        st.info(f"Ingen data för energi vs. duration tillgänglig för {area_filter_disp}.")
        return go.Figure()

    # Filtrera bort negativa eller noll-durationer för en meningsfull scatter plot om de inte redan hanterats
    scatter_data = filtered_df[(filtered_df['Varaktighet (minuter)'] > 0) & (filtered_df['Debiterad Energi (kWh)'] >= 0)].copy()
    
    if scatter_data.empty:
        st.info(f"Ingen giltig data (positiv duration) för energi vs. duration tillgänglig för {area_filter_disp}.")
        return go.Figure()

    # Lägg till en kolumn för laddhastighet (kWh/h)
    # se till att inte dividera med noll om Varaktighet (minuter) kan vara 0 efter filtrering
    scatter_data['Laddhastighet (kW)'] = (scatter_data['Debiterad Energi (kWh)'] / (scatter_data['Varaktighet (minuter)'] / 60)).replace([np.inf, -np.inf], np.nan)


    fig = px.scatter(scatter_data, 
                     x='Varaktighet (minuter)', 
                     y='Debiterad Energi (kWh)', 
                     title=f'Debiterad Energi vs. Laddningsduration ({area_filter_disp})',
                     color='Laddhastighet (kW)', # Färglägg punkter efter laddhastighet
                     color_continuous_scale=px.colors.sequential.Plasma,
                     hover_data=['ChargePoint', 'Laddhastighet (kW)'])
    fig.update_layout(
        xaxis_title="Laddningsduration (minuter)",
        yaxis_title="Debiterad Energi (kWh)",
        height=600
    )
    return fig

# --- PDF Generation ---
def fig_to_base64(fig):
    """Converts a Plotly figure to a base64 encoded image for PDF embedding."""
    try:
        img_bytes = fig.to_image(format="png", scale=2) # Öka scale för bättre upplösning
        buffered = io.BytesIO(img_bytes)
        # Använd PIL för att potentiellt optimera eller säkerställa korrekt format
        pil_img = PILImage.open(buffered)
        img_byte_arr = io.BytesIO()
        pil_img.save(img_byte_arr, format='PNG') # Spara som PNG
        img_byte_arr = img_byte_arr.getvalue()
        return base64.b64encode(img_byte_arr).decode()
    except Exception as e:
        st.error(f"Error converting figure to image: {e}")
        # Skapa en platshållarbild om konvertering misslyckas
        placeholder_img = PILImage.new('RGB', (500, 300), color = 'grey')
        img_byte_arr = io.BytesIO()
        placeholder_img.save(img_byte_arr, format='PNG')
        img_byte_arr = img_byte_arr.getvalue()
        return base64.b64encode(img_byte_arr).decode()


def generate_pdf(metrics, figures, area_filter_disp):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=inch/2, leftMargin=inch/2, topMargin=inch/2, bottomMargin=inch/2)
    styles = getSampleStyleSheet()
    
    # Anpassad stil för rubriker
    title_style = ParagraphStyle('TitleStyle', parent=styles['h1'], fontSize=18, spaceAfter=16, textColor=colors.HexColor("#2c3e50"))
    header_style = ParagraphStyle('HeaderStyle', parent=styles['h2'], fontSize=14, spaceAfter=10, textColor=colors.HexColor("#2c3e50"))
    body_style = styles["BodyText"]
    
    story = []

    # Titel
    story.append(Paragraph(f"Rapport för Laddstolpar - {area_filter_disp}", title_style))
    story.append(Paragraph(f"Genererad: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
    story.append(Spacer(1, 0.25 * inch))

    # Metrics
    if metrics:
        story.append(Paragraph("Nyckeltal", header_style))
        metrics_data = [["Mått", "Värde"]]
        for key, value in metrics.items():
            metrics_data.append([Paragraph(str(key), body_style), Paragraph(str(value), body_style)])
        
        # Tabell för nyckeltal
        table_metrics = Table(metrics_data, colWidths=[2.5*inch, 3.5*inch])
        table_metrics.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#4a6b82")),
            ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
            ('ALIGN',(0,0),(-1,-1),'LEFT'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 12),
            ('BACKGROUND',(0,1),(-1,-1),colors.HexColor("#d0dce4")),
            ('GRID',(0,0),(-1,-1),1,colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
        ]))
        story.append(table_metrics)
        story.append(Spacer(1, 0.25 * inch))

    # Figurer
    for fig_title, fig_obj in figures.items():
        if fig_obj is not None and not (isinstance(fig_obj, go.Figure) and not fig_obj.data): # Kontrollera om figuren har data
            story.append(Paragraph(fig_title, header_style))
            try:
                # Konvertera Plotly-figur till bild för ReportLab
                # Spara till temporär fil för Image-objektet
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
                    fig_obj.write_image(tmpfile.name, scale=2) # Öka scale för bättre upplösning
                    img = Image(tmpfile.name, width=7*inch, height=4*inch) # Justera storlek efter behov
                    img.hAlign = 'CENTER'
                    story.append(img)
            except Exception as e:
                story.append(Paragraph(f"Kunde inte rendera graf: {fig_title}. Fel: {e}", body_style))
            story.append(Spacer(1, 0.25 * inch))
        else:
            story.append(Paragraph(f"Ingen data tillgänglig för graf: {fig_title}", body_style))
            story.append(Spacer(1, 0.25 * inch))
            
    doc.build(story)
    buffer.seek(0)
    return buffer

# --- Streamlit App Layout ---
st.title("⚡ Dashboard för Laddstolpar")

# File uploader in the sidebar
st.sidebar.header("Ladda upp Filer")
uploaded_files = st.sidebar.file_uploader("Välj Excel (.xlsx) eller CSV (.csv) filer för laddsessioner", 
                                          type=["xlsx", "csv"], 
                                          accept_multiple_files=True, 
                                          help="Ladda upp en eller flera filer med data över laddsessioner. Förväntat format har kolumner som 'ChargePoint', 'Starttid', 'Debiterad Energi (kWh)', etc.")

if uploaded_files:
    try:
        # Ladda och bearbeta data
        df_combined = load_data(uploaded_files)

        if df_combined.empty:
            st.warning("Ingen data kunde laddas från de uppladdade filerna, eller all data filtrerades bort under rensning. Kontrollera filformat och innehåll.")
        else:
            st.success(f"{len(df_combined)} laddsessioner har laddats och bearbetats.")

            # Sidebar for filters
            st.sidebar.header("Filter")
            
            # Area filter - default to "All"
            # Se till att 'ChargePoint' finns och har giltiga värden
            area_options = ["All"]
            if 'ChargePoint' in df_combined.columns and df_combined['ChargePoint'].nunique() > 0:
                area_options.extend(sorted(df_combined['ChargePoint'].astype(str).unique().tolist()))
            
            selected_area = st.sidebar.selectbox("Välj Område/ChargePoint:", 
                                                 options=area_options,
                                                 index=0, # Default till "All"
                                                 help="Filtrera data baserat på specifikt område eller laddpunkt.")

            # Filter DataFrame based on selected_area
            if selected_area == "All":
                filtered_df = df_combined
                area_filter_display_name = "Alla Områden"
            else:
                filtered_df = df_combined[df_combined['ChargePoint'] == selected_area]
                area_filter_display_name = selected_area
            
            if filtered_df.empty and selected_area != "All":
                st.warning(f"Ingen data hittades för det valda området: {selected_area}. Visar data för alla områden istället eller ingen data om huvud-DataFrame är tom.")
                # filtered_df = df_combined # Återgå till all data om filtrering ger tomt resultat
                # area_filter_display_name = "Alla Områden (inget för val)"


            # Main content area with tabs
            st.header(f"Analyser för: {area_filter_display_name}")
            
            tab1, tab2, tab3, tab4 = st.tabs(["📊 Översikt & Nyckeltal", "🕒 Timvis Användning", "⚡ Energiförbrukning", "⏱️ Laddningsdetaljer"])

            figures_for_pdf = {} # Samla figurer för PDF-export

            with tab1:
                st.subheader("Nyckeltal")
                if not filtered_df.empty:
                    total_sessions = filtered_df.shape[0]
                    total_energy = filtered_df['Debiterad Energi (kWh)'].sum() if 'Debiterad Energi (kWh)' in filtered_df.columns else 0
                    avg_energy_per_session = filtered_df['Debiterad Energi (kWh)'].mean() if 'Debiterad Energi (kWh)' in filtered_df.columns and total_sessions > 0 else 0
                    avg_duration = filtered_df['Varaktighet (minuter)'].mean() if 'Varaktighet (minuter)' in filtered_df.columns and total_sessions > 0 else 0
                    unique_chargepoints = filtered_df['ChargePoint'].nunique() if 'ChargePoint' in filtered_df.columns else 0

                    metrics_summary = {
                        "Antal Laddsessioner": f"{total_sessions}",
                        "Total Debiterad Energi": f"{total_energy:.2f} kWh",
                        "Genomsnittlig Energi/Session": f"{avg_energy_per_session:.2f} kWh",
                        "Genomsnittlig Laddtid": f"{avg_duration:.2f} minuter",
                        "Antal Unika Laddpunkter (i urval)": f"{unique_chargepoints}"
                    }
                    
                    # Display metrics in columns
                    cols = st.columns(len(metrics_summary))
                    for i, (metric_name, metric_value) in enumerate(metrics_summary.items()):
                        cols[i].metric(metric_name, metric_value)
                    
                    figures_for_pdf["Nyckeltal (sammanfattning)"] = metrics_summary # Lägg till som dictionary, hanteras i PDF-generatorn

                else:
                    st.info("Ingen data att visa nyckeltal för efter filtrering.")
                
                st.subheader("Energiförbrukning per Laddpunkt")
                if not filtered_df.empty and 'ChargePoint' in filtered_df.columns and 'Debiterad Energi (kWh)' in filtered_df.columns:
                    energy_by_chargepoint = filtered_df.groupby('ChargePoint')['Debiterad Energi (kWh)'].sum().sort_values(ascending=False).reset_index()
                    if not energy_by_chargepoint.empty:
                        fig_energy_cp = px.bar(energy_by_chargepoint.head(15), x='ChargePoint', y='Debiterad Energi (kWh)', 
                                               title="Topp 15 Laddpunkter efter Energiförbrukning",
                                               labels={'ChargePoint': 'Laddpunkt', 'Debiterad Energi (kWh)': 'Total Debiterad Energi (kWh)'})
                        fig_energy_cp.update_layout(height=500)
                        st.plotly_chart(fig_energy_cp, use_container_width=True)
                        figures_for_pdf["Energiförbrukning per Laddpunkt"] = fig_energy_cp
                    else:
                        st.info("Kunde inte aggregera energiförbrukning per laddpunkt.")
                else:
                    st.info("Saknar data för 'ChargePoint' eller 'Debiterad Energi (kWh)' för att visa denna graf.")


            with tab2:
                st.subheader("Timvis Användning (Värmekarta)")
                if not filtered_df.empty:
                    fig_heatmap = plot_hourly_heatmap(filtered_df, selected_area if selected_area != "All" else "All") # Skicka med 'All' om det är valt
                    if fig_heatmap.data: # Kontrollera om figuren har data
                         st.plotly_chart(fig_heatmap, use_container_width=True)
                         figures_for_pdf["Timvis Användning (Värmekarta)"] = fig_heatmap
                    else:
                         st.info(f"Kunde inte generera värmekarta för {area_filter_display_name}.")
                else:
                    st.info("Ingen data att visa värmekarta för efter filtrering.")

            with tab3:
                st.subheader("Energiförbrukningstrender")
                if not filtered_df.empty:
                    fig_energy_trends = plot_energy_consumption_trends(filtered_df, selected_area if selected_area != "All" else "All")
                    if fig_energy_trends.data:
                        st.plotly_chart(fig_energy_trends, use_container_width=True)
                        figures_for_pdf["Energiförbrukningstrender"] = fig_energy_trends
                    else:
                        st.info(f"Kunde inte generera energitrender för {area_filter_display_name}.")

                else:
                    st.info("Ingen data att visa energitrender för efter filtrering.")
                
                st.subheader("Energi vs. Laddningsduration")
                if not filtered_df.empty:
                    fig_energy_duration = plot_energy_vs_duration(filtered_df, selected_area if selected_area != "All" else "All")
                    if fig_energy_duration.data:
                        st.plotly_chart(fig_energy_duration, use_container_width=True)
                        figures_for_pdf["Energi vs. Laddningsduration"] = fig_energy_duration
                    else:
                        st.info(f"Kunde inte generera graf för energi vs duration för {area_filter_display_name}.")
                else:
                    st.info("Ingen data att visa energi vs duration för efter filtrering.")

            with tab4:
                st.subheader("Distribution av Laddningsduration")
                if not filtered_df.empty:
                    fig_duration_dist = plot_charging_duration_distribution(filtered_df, selected_area if selected_area != "All" else "All")
                    if fig_duration_dist.data:
                        st.plotly_chart(fig_duration_dist, use_container_width=True)
                        figures_for_pdf["Distribution av Laddningsduration"] = fig_duration_dist
                    else:
                        st.info(f"Kunde inte generera graf för laddningsduration för {area_filter_display_name}.")
                else:
                    st.info("Ingen data att visa laddningsduration för efter filtrering.")

                st.subheader("Distribution av Start- och Slut-SoC")
                if not filtered_df.empty:
                    fig_soc_dist = plot_soc_distribution(filtered_df, selected_area if selected_area != "All" else "All")
                    # Kontrollera om fig_soc_dist innehåller några spår (traces)
                    if fig_soc_dist.data and any(trace for trace in fig_soc_dist.data):
                        st.plotly_chart(fig_soc_dist, use_container_width=True)
                        figures_for_pdf["Distribution av Start- och Slut-SoC"] = fig_soc_dist
                    else:
                        st.info(f"Kunde inte generera SoC-distribution för {area_filter_display_name} (möjligen saknas SoC-data).")
                else:
                    st.info("Ingen data att visa SoC-distribution för efter filtrering.")
            
            # Generate PDF button
            st.sidebar.header("Exportera Rapport")
            if st.sidebar.button("Generera PDF Rapport"):
                # Samla ihop de faktiska nyckeltalen från metrics_summary för PDF
                pdf_metrics = metrics_summary if 'metrics_summary' in locals() and not filtered_df.empty else {"Info": "Ingen data tillgänglig för nyckeltal."}
                
                # Skapa en ny dictionary för PDF-figurer som bara innehåller giltiga Plotly-figurer
                valid_figures_for_pdf = {}
                for title, fig_or_data in figures_for_pdf.items():
                    if isinstance(fig_or_data, go.Figure) and fig_or_data.data : # Kontrollera om det är en figur med data
                        valid_figures_for_pdf[title] = fig_or_data
                    elif isinstance(fig_or_data, dict): # Hantera nyckeltal som dictionary
                         valid_figures_for_pdf[title] = fig_or_data


                if not valid_figures_for_pdf and not (isinstance(pdf_metrics, dict) and pdf_metrics.get("Info") is None):
                     st.sidebar.warning("Inga grafer eller tillräckligt med data för att generera PDF.")
                else:
                    with st.spinner("Genererar PDF rapport..."):
                        try:
                            pdf_buffer = generate_pdf(pdf_metrics, valid_figures_for_pdf, area_filter_display_name)
                            
                            b64_pdf = base64.b64encode(pdf_buffer.read()).decode()
                            href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="Laddrapport_{area_filter_display_name.replace(" ", "_")}_{datetime.now().strftime("%Y%m%d")}.pdf">Ladda ner PDF Rapport</a>'
                            st.sidebar.markdown(href, unsafe_allow_html=True)
                            st.sidebar.success("PDF genererad!")
                        except Exception as pdf_e:
                            st.sidebar.error(f"Kunde inte generera PDF: {pdf_e}")
                            st.sidebar.exception(pdf_e)
    
    except Exception as e:
        st.error(f"Ett oväntat fel uppstod vid bearbetning av data: {e}")
        st.exception(e)
else:
    st.info("Vänligen ladda upp en eller flera datafiler (Excel eller CSV) för att påbörja analysen.")
    
    st.header("Förhandsgranskning av Dashboard")
    st.markdown("""
    Denna dashboard hjälper dig att analysera prestandan för laddstolpar med:
    
    1.  **Nyckeltal** - Antal sessioner, total energi, genomsnittlig laddtid per område, etc.
    2.  **Användningsanalys** - Värmekarta som visar mönster för när uttagen används mest.
    3.  **Energiförbrukning** - Detaljerade grafer över energianvändning över tid och per laddpunkt.
    4.  **Laddningsdetaljer** - Distribution av laddtider och State of Charge (SoC).
    
    Du kan filtrera data baserat på specifika områden/laddpunkter och generera PDF-rapporter.
    """)
    st.markdown("---")
    st.caption(f"Laddstolpsanalys v1.1 - Senast uppdaterad: {datetime.now().strftime('%Y-%m-%d')}")