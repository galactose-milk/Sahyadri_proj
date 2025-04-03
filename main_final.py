import streamlit as st
import pandas as pd
import plotly.express as px
import ezodf
from datetime import datetime

def parse_date(date_str):
    """Attempts to parse a date from multiple formats."""
    date_formats = ['%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y']
    for fmt in date_formats:
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None  # If all formats fail

def load_rejection_trend_data(sheet):
    """Extracts rejection percentage trend data."""
    rejection_trend_data = []
    for row in range(16, 47):  # Rows B17 to B47
        date_str = str(sheet[row, 1].value).strip() if sheet[row, 1].value else None
        rejection_value = sheet[row, 25].value  # Column Z (index 25)
        
        if not date_str:
            continue
        
        date = parse_date(date_str)
        if not date:
            continue
        
        try:
            rejection_percentage = float(rejection_value) if rejection_value not in ["#DIV/0!", None] else None
        except ValueError:
            continue
        
        rejection_trend_data.append({'Date': date, 'Rejection %': rejection_percentage})
    
    if not rejection_trend_data:
        return None
    
    df = pd.DataFrame(rejection_trend_data).sort_values(by='Date')
    return df

def load_rejection_breakdown_data(sheet):
    """Extracts rejection breakdown by type."""
    total_sheets = float(sheet[49, 2].value or 0)
    rejection_columns = range(3, 24)  # Columns C to W
    rejection_types = [
        'Layer Open', 'Water Mark', 'Extra Material', 'Bend', 'Edge Damage', 
        'TWM', 'VC', 'TC', 'Cutting Mist', 'Side Damage', 
        'Corner Damage', 'LT', 'FT', 'Surf Defect', 'Hole Damage', 
        'Temp Damage', 'Rolling Particle', 'TV', 'GSD', 'Brittle', 
        'Lab Sheet'
    ]
    rejection_data = []
    total_rejection_sheets = 0
    for col, rej_type in zip(rejection_columns, rejection_types):
        rejection_sheets = float(sheet[49, col].value or 0)
        total_rejection_sheets += rejection_sheets
        rejection_percentage = (rejection_sheets / total_sheets) * 100 if total_sheets else 0
        rejection_data.append({'Rejection Type': rej_type, 'Rejection Sheets': rejection_sheets, 'Rejection Percentage': rejection_percentage})
    df = pd.DataFrame(rejection_data).sort_values('Rejection Percentage', ascending=False)
    total_rejection_percentage = (total_rejection_sheets / total_sheets) * 100 if total_sheets else 0
    return df, total_sheets, total_rejection_sheets, total_rejection_percentage

def main():
    st.title('📊 Stamping Rejection Analysis')
    uploaded_file = st.file_uploader("📂 Upload an ODS file", type=['ods'])
    if not uploaded_file:
        st.warning("⚠️ Please upload a file to proceed.")
        return
    
    with open("temp.ods", "wb") as f:
        f.write(uploaded_file.read())
    
    doc = ezodf.opendoc("temp.ods")
    sheet = next((s for s in doc.sheets if s.name == 'Stamping Rej'), None)
    if not sheet:
        st.error("❌ Sheet 'Stamping Rej' not found!")
        return
    
    rejection_trend_df = load_rejection_trend_data(sheet)
    rejection_breakdown_result = load_rejection_breakdown_data(sheet)
    
    if rejection_trend_df is not None:
        st.header('📈 Rejection Percentage Trend Over Time')
        fig = px.line(rejection_trend_df, x='Date', y='Rejection %',
                      title='Rejection Percentage Trend', labels={'Rejection %': 'Rejection Percentage (%)'})
        st.plotly_chart(fig)
        st.dataframe(rejection_trend_df.style.format({'Date': lambda x: x.strftime('%d/%m/%Y'), 'Rejection %': '{:.2f}%'}))
    
    if rejection_breakdown_result:
        rejection_df, total_sheets, total_rejection_sheets, total_rejection_percentage = rejection_breakdown_result
        
        st.header('🔍 Rejection Breakdown by Type')
        top_5_df = rejection_df.head()
        fig = px.bar(top_5_df, x='Rejection Type', y='Rejection Percentage',
                     title=f'Top 5 Rejection Types (Total Rejection: {total_rejection_percentage:.2f}%)',
                     labels={'Rejection Percentage': 'Rejection Percentage (%)'}, color='Rejection Type')
        st.plotly_chart(fig)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Sheets", f"{total_sheets:,.0f}")
        with col2:
            st.metric("Total Rejection Sheets", f"{total_rejection_sheets:,.0f}")
        with col3:
            st.metric("Total Rejection Percentage", f"{total_rejection_percentage:.2f}%")
        
        st.header('📋 Detailed Rejection Data')
        st.dataframe(rejection_df.style.format({'Rejection Sheets': '{:.2f}', 'Rejection Percentage': '{:.2f}%'}))

if __name__ == '__main__':
    main()
