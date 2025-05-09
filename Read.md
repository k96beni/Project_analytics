Charging Outlets Analytics Dashboard
This Streamlit application helps you analyze charging outlet performance data and create insightful visualizations.

Features
Data Upload: Upload Sessions.xlsx and Overview.xlsx files containing charging session data
Interactive Dashboard: Visualize key metrics about charging outlets
Filtering: Filter data by specific areas
PDF Export: Generate downloadable PDF reports with visualizations
Dashboard Views
Key Metrics
Number of outlets per area
kWh/month per area
Total sessions and energy consumption
Utilization Analysis
Heatmap of outlet utilization (used outlet days / total outlet days)
Detailed utilization metrics table
Energy Consumption Analysis
kWh per outlet per month per area
Average kWh per session
Deployment on Hugging Face Spaces
Prerequisites
A Hugging Face account
Basic knowledge of Git
Deployment Steps
Create a new Space on Hugging Face:
Go to https://huggingface.co/spaces
Click "Create new Space"
Select "Streamlit" as the SDK
Name your Space and set visibility options
Clone the repository locally:
bash
git clone https://huggingface.co/spaces/YOUR-USERNAME/YOUR-SPACE-NAME
Add the application files to the repository:
Copy app.py to the root of the repository
Copy requirements.txt to the root of the repository
Commit and push your changes:
bash
git add .
git commit -m "Initial application setup"
git push
Wait for the build to complete on Hugging Face Spaces
Your application is now live at https://huggingface.co/spaces/YOUR-USERNAME/YOUR-SPACE-NAME
Data Format
Sessions.xlsx
Contains charging session data with the following columns:

Områdeskod: Area code
Område: Area name
Uttag: Outlet identifier
Startad: Session start time
Avslutad: Session end time
Laddat (kWh): Energy consumed in kWh
Kostnad (exkl): Cost (excluding tax)
Valuta: Currency
Overview.xlsx
Contains area and outlet information with the following columns:

Namn: Name
Avtalskund: Contract customer
Synlighet: Visibility
Antal uttag: Number of outlets
Start fasta avgifter: Start of fixed fees
Fakturerbara ACAntal uttag: Billable AC number of outlets
Fakturerbara DCAntal uttag: Billable DC number of outlets
Kolumn1: Additional column
Usage
Upload both required files using the file uploaders in the sidebar
Use the dropdown in the sidebar to filter by specific areas
Navigate between tabs to view different dashboards
Generate a PDF report using the button in the sidebar
Requirements
streamlit==1.27.0
pandas==2.0.3
plotly==5.17.0
numpy==1.25.2
reportlab==4.0.4
pillow==10.0.1
openpyxl==3.1.2
