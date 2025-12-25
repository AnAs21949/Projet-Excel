import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import re
from io import BytesIO
import tempfile
import numpy as np


def extract_numeric(row, col_idx):
    """Safely extract numeric value from row"""
    if col_idx is None or col_idx >= len(row):
        return np.nan
    value = row[col_idx]
    if pd.isna(value):
        return np.nan
    try:
        return float(value)
    except:
        return np.nan


def detect_format_and_columns(data):
    """Detect file format and find column positions for TOTAL MATRICULE rows"""
    num_cols = len(data.columns)
    
    # Find a TOTAL MATRICULE row to detect column positions
    for index, row in data.iterrows():
        row_str = ' '.join(map(str, row.dropna().values))
        if 'TOTAL MATRICULE' in row_str:
            # Based on column count, determine format
            if num_cols <= 10:  # 7-column format (compact)
                return {
                    'Tps_Saisie': 1,
                    'Tps_Alloue': 2,
                    'Ecart': 3,
                    'Efficience_Tech': 4
                }
            else:  # 26-column format (expanded)
                return {
                    'Tps_Saisie': 13,
                    'Tps_Alloue': 16,
                    'Ecart': 18,
                    'Efficience_Tech': 25
                }
    
    # Default fallback
    return {
        'Tps_Saisie': 1,
        'Tps_Alloue': 2,
        'Ecart': 3,
        'Efficience_Tech': 4
    }


def extract_employee_data(file_path: str):
    """Extract employee efficiency data from X3 report - handles multiple formats"""
    data = pd.read_excel(file_path, header=None, engine='xlrd')
    
    # Auto-detect format and column positions
    col_map = detect_format_and_columns(data)
    
    employees = []
    current_employee = {}
    current_period = None

    for index, row in data.iterrows():
        row_str = ' '.join(map(str, row.dropna().values))

        # Match employee header: "Matricule : 7 001 - SADIK Amina"
        matricule_match = re.search(r'Matricule\s*:\s*(\d[\d\s]+\d)\s*-\s*(.+?)(?:\s+Période|$)', row_str)
        if matricule_match:
            current_employee = {
                'Matricule': matricule_match.group(1).replace(' ', ''),
                'Employee_Name': matricule_match.group(2).strip()
            }
            current_period = None

        # Match period in data rows (e.g., "2025/12")
        period_match = re.search(r'^(\d{4}/\d{2})\s', row_str)
        if period_match and 'TOTAL' not in row_str:
            current_period = period_match.group(1)

        # Match TOTAL MATRICULE row - extract totals using detected column positions
        total_match = re.search(r'TOTAL MATRICULE\s*:\s*(\d[\d\s]+\d)', row_str)
        if total_match and current_employee:
            employee_data = {
                **current_employee,
                'Period': current_period if current_period else 'N/A',
                'Tps_Saisie': extract_numeric(row, col_map['Tps_Saisie']),
                'Tps_Alloue': extract_numeric(row, col_map['Tps_Alloue']),
                'Ecart': extract_numeric(row, col_map['Ecart']),
                'Efficience_Tech': extract_numeric(row, col_map['Efficience_Tech'])
            }
            employees.append(employee_data)
            current_employee = {}

    return pd.DataFrame(employees)


# Configure page
st.set_page_config(
    page_title="Production Efficiency Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("Production Efficiency Dashboard")
st.markdown("Upload X3 production efficiency reports and get instant analysis")

# Sidebar
with st.sidebar:
    st.header("Upload Files")
    uploaded_files = st.file_uploader(
        "Choose Excel files (.xls)",
        type=['xls'],
        accept_multiple_files=True,
        help="Upload one or more X3 production efficiency reports"
    )

    st.markdown("---")
    st.markdown("### About")
    st.markdown("""
    This dashboard analyzes X3 production efficiency reports and provides:
    - Employee efficiency metrics
    - Visual analytics
    - Downloadable clean data
    - Period comparisons
    """)

# Main content
if not uploaded_files:
    st.info("Please upload one or more Excel files to get started")

    # instructions
    st.markdown("### Expected File Format")
    st.markdown("""
    Your Excel file should contain:
    - Employee matricule and names
    - Work periods (YYYY/MM)
    - Tps Saisie (actual hours)
    - Tps Alloué (allocated hours)
    - Ecart (variance)
    - Efficience Tech (efficiency %)
    """)

else:
    #  uploaded files
    all_data = []

    for uploaded_file in uploaded_files:
        with st.spinner(f'Processing {uploaded_file.name}...'):
            # Save uploaded file temporarily
            temp_dir = tempfile.gettempdir()
            temp_path = Path(temp_dir) / uploaded_file.name
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Extract data
            df = extract_employee_data(temp_path)
            df['Source_File'] = uploaded_file.name
            all_data.append(df)

    # Combine all data
    combined_df = pd.concat(all_data, ignore_index=True)

    # Display metrics
    st.markdown("## Key Metrics")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Total Employees", len(combined_df))

    with col2:
        avg_efficiency = combined_df['Efficience_Tech'].mean() * 100
        st.metric("Average Efficiency", f"{avg_efficiency:.1f}%")

    with col3:
        total_hours = combined_df['Tps_Saisie'].sum()
        st.metric("Total Hours Worked", f"{total_hours:.0f}h")

    with col4:
        periods = combined_df['Period'].unique()
        st.metric("Periods", len(periods))

    tab1, tab2, tab3, tab4 = st.tabs(["Charts", "Data Table", "Employee Details", "Download"])

    with tab1:
        st.markdown("### Efficiency Distribution")

        col1, col2 = st.columns(2)

        with col1:
            fig_hist = px.histogram(
                combined_df,
                x='Efficience_Tech',
                nbins=20,
                title='Efficiency Distribution',
                labels={'Efficience_Tech': 'Efficiency', 'count': 'Number of Employees'},
                color_discrete_sequence=['#1f77b4']
            )
            fig_hist.update_layout(showlegend=False)
            st.plotly_chart(fig_hist, use_container_width=True)

        with col2:
            top_10 = combined_df.nlargest(10, 'Efficience_Tech')
            fig_bar = px.bar(
                top_10,
                x='Efficience_Tech',
                y='Employee_Name',
                orientation='h',
                title='Top 10 Most Efficient Employees',
                labels={'Efficience_Tech': 'Efficiency', 'Employee_Name': 'Employee'},
                color='Efficience_Tech',
                color_continuous_scale='Greens'
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        # Period comparison (if multiple periods)
        if len(periods) > 1:
            st.markdown("### Period Comparison")
            period_stats = combined_df.groupby('Period').agg({
                'Efficience_Tech': 'mean',
                'Tps_Saisie': 'sum',
                'Matricule': 'count'
            }).reset_index()
            period_stats.columns = ['Period', 'Avg Efficiency', 'Total Hours', 'Employee Count']

            fig_period = go.Figure()
            fig_period.add_trace(go.Bar(
                x=period_stats['Period'],
                y=period_stats['Avg Efficiency'],
                name='Avg Efficiency',
                marker_color='lightblue'
            ))
            fig_period.update_layout(
                title='Average Efficiency by Period',
                xaxis_title='Period',
                yaxis_title='Efficiency',
                showlegend=False
            )
            st.plotly_chart(fig_period, use_container_width=True)

        st.markdown("### Hours Worked vs Efficiency")
        fig_scatter = px.scatter(
            combined_df,
            x='Tps_Saisie',
            y='Efficience_Tech',
            color='Period',
            hover_data=['Employee_Name', 'Matricule'],
            title='Hours Worked vs Efficiency',
            labels={'Tps_Saisie': 'Hours Worked', 'Efficience_Tech': 'Efficiency'}
        )
        st.plotly_chart(fig_scatter, use_container_width=True)

    with tab2:
        st.markdown("### Data Table")

        # filters
        col1, col2 = st.columns(2)

        with col1:
            period_filter = st.multiselect(
                "Filter by Period",
                options=combined_df['Period'].unique(),
                default=combined_df['Period'].unique()
            )

        with col2:
            min_efficiency = st.slider(
                "Minimum Efficiency",
                min_value=0.0,
                max_value=float(combined_df['Efficience_Tech'].max()),
                value=0.0
            )

        filtered_df = combined_df[
            (combined_df['Period'].isin(period_filter)) &
            (combined_df['Efficience_Tech'] >= min_efficiency)
        ]

        # Display table
        st.dataframe(
            filtered_df.style.format({
                'Tps_Saisie': '{:.2f}',
                'Tps_Alloue': '{:.2f}',
                'Ecart': '{:.2f}',
                'Efficience_Tech': '{:.2%}'
            }),
            use_container_width=True,
            height=400
        )

        st.markdown(f"**Showing {len(filtered_df)} of {len(combined_df)} employees**")

    with tab3:
        st.markdown("### Employee Search")

        # Search by name or matricule
        search_term = st.text_input("Search by name or matricule", "")

        if search_term:
            search_results = combined_df[
                combined_df['Employee_Name'].str.contains(search_term, case=False, na=False) |
                combined_df['Matricule'].str.contains(search_term, case=False, na=False)
            ]

            if len(search_results) > 0:
                for idx, row in search_results.iterrows():
                    with st.expander(f"{row['Matricule']} - {row['Employee_Name']}"):
                        col1, col2, col3 = st.columns(3)

                        with col1:
                            st.metric("Period", row['Period'])
                            st.metric("Hours Worked", f"{row['Tps_Saisie']:.2f}h")

                        with col2:
                            st.metric("Allocated Hours", f"{row['Tps_Alloue']:.2f}h")
                            st.metric("Variance", f"{row['Ecart']:.2f}h")

                        with col3:
                            efficiency_pct = row['Efficience_Tech'] * 100
                            st.metric("Efficiency", f"{efficiency_pct:.1f}%")

                            if row['Efficience_Tech'] >= 1.0:
                                st.success("Above target")
                            elif row['Efficience_Tech'] >= 0.8:
                                st.warning("Near target")
                            else:
                                st.error("Error: Below target")
            else:
                st.warning("No employees found matching your search")

    with tab4:
        st.markdown("### Download Cleaned Data")

        col1, col2 = st.columns(2)

        with col1:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False, sheet_name='Employee_Efficiency')

            st.download_button(
                label="Download as Excel",
                data=output.getvalue(),
                file_name="cleaned_efficiency_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col2:
            csv = combined_df.to_csv(index=False)
            st.download_button(
                label="Download as CSV",
                data=csv,
                file_name="cleaned_efficiency_data.csv",
                mime="text/csv"
            )

        st.markdown("### Summary Statistics")
        summary = combined_df.describe()
        st.dataframe(summary)