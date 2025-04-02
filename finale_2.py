import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objs as go
import ezodf

def load_rejection_data():
    """Load rejection data from the OpenDocument Spreadsheet file."""
    # Open the .ods file
    doc = ezodf.opendoc('/home/galactose/Downloads/sahyadri_march.ods')
    
    # Select the 'Stamping Rej' sheet
    sheet = None
    for s in doc.sheets:
        if s.name == 'Stamping Rej':
            sheet = s
            break
    
    if not sheet:
        st.error("Sheet 'Stamping Rej' not found in the document.")
        return None
    
    # Total sheets produced (row 50, column C - remember 0-based indexing)
    total_sheets = float(sheet[49, 2].value or 0)
    
    # Rejection columns and types
    rejection_columns = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]
    rejection_types = [
        'Layer Open', 'Water Mark', 'Extra Material', 'Bend', 'Edge Damage', 
        'TWM', 'VC', 'TC', 'Cutting Mist', 'Side Damage', 
        'Corner Damage', 'LT', 'FT', 'Surf Defect', 'Hole Damge', 
        'Temp Damage', 'Rolling Particle', 'TV', 'GSD', 'Brittle', 
        'Lab Sheet'
    ]
    # Calculate rejection percentages
    rejection_data = []
    total_rejection_sheets = 0
    for col, rej_type in zip(rejection_columns, rejection_types):
        rejection_sheets = float(sheet[49, col].value or 0)
        total_rejection_sheets += rejection_sheets
        rejection_percentage = (rejection_sheets / total_sheets) * 100
        rejection_data.append({
            'Rejection Type': rej_type,
            'Rejection Sheets': rejection_sheets,
            'Rejection Percentage': rejection_percentage
        })6.10%
    # Convert to DataFrame and sort
    rejection_df = pd.DataFrame(rejection_data)
    rejection_df = rejection_df.sort_values('Rejection Percentage', ascending=False)
    
    return rejection_df, total_sheets, total_rejection_sheets, total_rejection_percentage

def main():
    st.title('Stamping Rejection Analysis')
    
    # Load data
    result = load_rejection_data()
    if result is None:
        return
    
    rejection_df, total_sheets, total_rejection_sheets, total_rejection_percentage = result
    
    # Top 5 Rejections
    st.header('Top 5 Rejection Types')
    top_5_df = rejection_df.head()
    
    # Bar Chart using Plotly
    fig = px.bar(
        top_5_df, 
        x='Rejection Type', 
        y='Rejection Percentage',
        title=f'Top 5 Rejection Types (Total Rejection: {total_rejection_percentage:.2f}%)',
        labels={'Rejection Percentage': 'Rejection Percentage (%)', 'Rejection Type': 'Rejection Type'},
        color='Rejection Type'
    )
    
    # Customize the layout
    fig.update_layout(
        xaxis_title='Rejection Type',
        yaxis_title='Rejection Percentage (%)',
        height=500,
        width=800,
        annotations=[
            dict(
                x=1.05,
                y=1.1,
                xref='paper',
                yref='paper',
                text=f'Total Rejection: {total_rejection_percentage:.2f}%',
                showarrow=False,
                font=dict(size=14, color='red')
            )
        ]
    )
    
    # Display the chart
    st.plotly_chart(fig)
    
    # Summary Statistics
    st.header('Rejection Summary')
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Sheets", f"{total_sheets:,.0f}")
    with col2:
        st.metric("Total Rejection Sheets", f"{total_rejection_sheets:,.0f}")
    with col3:
        st.metric("Total Rejection Percentage", f"{total_rejection_percentage:.2f}%")
    
    # Detailed Table
    st.header('Detailed Rejection Analysis')
    st.dataframe(rejection_df[['Rejection Type', 'Rejection Sheets', 'Rejection Percentage']].style.format({
        'Rejection Sheets': '{:.2f}',
        'Rejection Percentage': '{:.2f}%'
    }))

if __name__ == '__main__':
    main()

# requirements.txt content:
# streamlit
# pandas
# plotly
# ezodf
# lxml