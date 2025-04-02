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

def load_rejection_data():
    """Load rejection data from the OpenDocument Spreadsheet file."""
    doc = ezodf.opendoc('/home/galactose/Downloads/sahyadri_march.ods')

    # Select the 'Stamping Rej' sheet
    sheet = next((s for s in doc.sheets if s.name == 'Stamping Rej'), None)
    if not sheet:
        st.error("❌ Sheet 'Stamping Rej' not found!")
        return None
    
    # Collect data
    rejection_trend_data = []
    
    for row in range(16, 47):  # Python index: B17 (row 16) to B47 (row 46)
        date_str = str(sheet[row, 1].value).strip() if sheet[row, 1].value else None
        rejection_value = sheet[row, 25].value  # Z column (column index 25)

        # Handle missing date
        if not date_str:
            st.warning(f"⚠️ Skipping row {row + 1}: No date value.")
            continue

        # Parse date
        date = parse_date(date_str)
        if not date:
            st.error(f"❌ Invalid date format in row {row + 1}: '{date_str}'")
            continue

        # Handle invalid rejection values
        try:
            rejection_percentage = float(rejection_value) if rejection_value not in ["#DIV/0!", None] else None
        except ValueError:
            st.warning(f"⚠️ Skipping row {row + 1}: Invalid rejection value '{rejection_value}'")
            continue

        # Add to data list
        rejection_trend_data.append({'Date': date, 'Rejection %': rejection_percentage})

    if not rejection_trend_data:
        st.error("❌ No valid data found.")
        return None

    # Convert to DataFrame and sort by date
    rejection_trend_df = pd.DataFrame(rejection_trend_data).sort_values(by='Date')
    return rejection_trend_df

def main():
    st.title('📊 Stamping Rejection Trend Analysis')

    rejection_trend_df = load_rejection_data()
    if rejection_trend_df is None or rejection_trend_df.empty:
        return

    # Line Chart of Rejection % over Time
    fig = px.line(
        rejection_trend_df, 
        x='Date', 
        y='Rejection %',
        title='📈 Rejection Percentage Trend Over Time',
        labels={'Rejection %': 'Rejection Percentage (%)', 'Date': 'Date'}
    )

    fig.update_layout(
        xaxis_title='Date',
        yaxis_title='Rejection Percentage (%)',
        height=500,
        width=800
    )

    st.plotly_chart(fig)

    # Show Data
    st.header('📋 Rejection Trend Data')
    st.dataframe(rejection_trend_df.style.format({
        'Date': lambda x: x.strftime('%d/%m/%Y'),
        'Rejection %': lambda x: f'{x:.2f}%' if x is not None else '-'
    }))

if __name__ == '__main__':
    main()


#conclusion instead of keeping values at last blank what can we do is we fill it with the average value of the graph or the first and last non missing values and thjeir average so that the graph remains continous....