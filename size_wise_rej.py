import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objs as go
import ezodf
from datetime import datetime
import anthropic
import re

def load_raw_sheet_data():
    """
    Load raw data from the OpenDocument Spreadsheet file.
    Capture entire sheet content for AI preprocessing.
    """
    # Open the .ods file
    doc = ezodf.opendoc('/home/galactose/Downloads/sahyadri_march.ods')
    
    # Select the 'Size wise Rej' sheet
    sheet = None
    for s in doc.sheets:
        if s.name == 'Size wise Rej':
            sheet = s
            break
    
    if not sheet:
        st.error("Sheet 'Size wise Rej' not found in the document.")
        return None
    
    # Prepare raw data extraction
    raw_data = []
    headers = []
    
    # Extract headers
    for col in range(sheet.ncols()):
        headers.append(str(sheet[0, col].value).strip() if sheet[0, col].value else f'Column_{col}')
    
    # Extract all rows with their raw values
    for row in range(1, sheet.nrows()):
        row_data = {}
        for col, header in enumerate(headers):
            cell_value = sheet[row, col].value
            row_data[header] = str(cell_value).strip() if cell_value is not None else ''
        raw_data.append(row_data)
    
    return raw_data, headers

def preprocess_data_with_ai(raw_data, headers):
    """
    Use Claude to preprocess and clean the raw data
    """
    # Convert raw data to a string representation
    data_str = "Headers: " + ", ".join(headers) + "\n\n"
    data_str += "\n".join([str(row) for row in raw_data[:50]])  # Limit to first 50 rows to avoid context limits
    
    # Prepare the client
    client = anthropic.Anthropic()
    
    # Prepare the prompt for data preprocessing
    prompt = f"""I have an Excel sheet with potentially inconsistent data. Analyze the following data:

{data_str}

Please help me:
1. Identify the correct columns for date, thickness, and rejection percentage
2.plot a graph if you can
3. Suggest a standardized approach to parse dates (what formats are present?)
4. Provide a strategy to clean and convert date, thickness, and rejection percentage columns
5. Handle any inconsistencies or missing values
6. Recommend how to convert these to numerical/datetime values

Respond with:
- Plot a graph between rejection percentage and thickness
- Exact column names for date, thickness, and rejection percentage
- Proposed parsing strategy
- Any special handling needed
- Sample code or transformation steps"""
    
    try:
        # Send request to Claude
        response = client.messages.create(
            model="claude-3-opus-20240229",
            max_tokens=1000,
            messages=[
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )
        
        # Return the analysis text
        return response.content[0].text
    except Exception as e:
        return f"Error in LLM preprocessing analysis: {str(e)}"

def process_data_with_ai_guidance(raw_data, ai_guidance):
    """
    Process the data based on AI-generated guidance
    """
    # Parse the AI guidance to extract key information
    def extract_column_name(guidance, column_type):
        # Use regex to find column names in the guidance
        pattern = rf"{column_type}\s*column\s*[:]?\s*['\"](.*?)['\"]"
        match = re.search(pattern, guidance, re.IGNORECASE)
        return match.group(1) if match else None
    
    # Extract column names
    date_column = extract_column_name(ai_guidance, 'date')
    thickness_column = extract_column_name(ai_guidance, 'thickness')
    rejection_column = extract_column_name(ai_guidance, 'rejection')
    
    # If columns not found, fallback to manual mapping
    if not all([date_column, thickness_column, rejection_column]):
        st.warning("Automatic column detection failed. Manual mapping may be required.")
        return None
    
    # Prepare processed data
    processed_data = []
    
    # Date parsing function
    def parse_date(date_str):
        # List of potential date formats to try
        date_formats = [
            '%d-%m-%Y', '%m-%d-%Y', '%Y-%m-%d',  # Common formats
            '%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d',  # Slash-separated
            '%B %d, %Y', '%d %B %Y',  # Full month name
            '%b %d, %Y', '%d %b %Y'   # Abbreviated month name
        ]
        
        # Try parsing with different formats
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except (ValueError, TypeError):
                continue
        
        # If no format works, return None
        return None
    
    # Process each row
    for row in raw_data:
        try:
            # Parse date
            date = parse_date(row.get(date_column, ''))
            
            # Convert thickness to float
            thickness = float(re.sub(r'[^\d.]', '', str(row.get(thickness_column, ''))))
            
            # Convert rejection percentage to float
            rejection_str = str(row.get(rejection_column, '')).replace(',', '.')
            rejection_percentage = float(re.sub(r'[^\d.]', '', rejection_str))
            
            # Only add row if all critical values are present
            if date and thickness and rejection_percentage is not None:
                processed_data.append({
                    'Date': date,
                    'Thickness': thickness,
                    'Rejection Percentage': rejection_percentage
                })
        
        except (ValueError, TypeError) as e:
            # Log parsing errors
            st.warning(f"Error processing row: {row}. Error: {e}")
    
    return pd.DataFrame(processed_data)

def main():
    st.title('AI-Powered Excel Data Processing and Analysis')
    
    # Load raw data
    raw_data = load_raw_sheet_data()
    
    if raw_data is None:
        st.error("No data could be loaded.")
        return
    
    # Unpack raw data
    raw_rows, headers = raw_data
    
    # Create tabs for different views
    tab1, tab2, tab3 = st.tabs(['AI Preprocessing Guidance', 'Data Processing', 'Data Visualization'])
    
    with tab1:
        st.header('AI Data Preprocessing Analysis')
        # Get AI guidance on data preprocessing
        ai_preprocessing_guidance = preprocess_data_with_ai(raw_rows, headers)
        st.write(ai_preprocessing_guidance)
    
    with tab2:
        st.header('Processed Data')
        # Process data based on AI guidance
        processed_df = process_data_with_ai_guidance(raw_rows, ai_preprocessing_guidance)
        
        if processed_df is not None and not processed_df.empty:
            st.dataframe(processed_df)
            
            # Descriptive Statistics
            st.header('Descriptive Statistics')
            thickness_stats = processed_df.groupby('Thickness')['Rejection Percentage'].agg([
                ('Mean', 'mean'),
                ('Min', 'min'),
                ('Max', 'max'),
                ('Standard Deviation', 'std')
            ]).reset_index()
            
            st.dataframe(thickness_stats.style.format({
                'Mean': '{:.2f}',
                'Min': '{:.2f}',
                'Max': '{:.2f}',
                'Standard Deviation': '{:.2f}'
            }))
        else:
            st.error("Could not process data. Please review AI guidance and manually map columns.")
    
    with tab3:
        if processed_df is not None and not processed_df.empty:
            # 3D Scatter Plot
            fig_3d = px.scatter_3d(
                processed_df, 
                x='Date', 
                y='Thickness', 
                z='Rejection Percentage',
                color='Thickness',
                title='Rejection Percentage by Date and Thickness',
                labels={
                    'Date': 'Date', 
                    'Thickness': 'Thickness (mm)', 
                    'Rejection Percentage': 'Rejection Percentage (%)'
                }
            )
            
            # Customize 3D plot layout
            fig_3d.update_layout(
                scene=dict(
                    xaxis_title='Date',
                    yaxis_title='Thickness (mm)',
                    zaxis_title='Rejection Percentage (%)'
                ),
                height=600,
                width=900
            )
            
            st.plotly_chart(fig_3d)
            
            # Line plot of rejection percentage over time
            grouped_df = processed_df.groupby(['Date', 'Thickness'])['Rejection Percentage'].mean().reset_index()
            pivot_df = grouped_df.pivot(index='Date', columns='Thickness', values='Rejection Percentage')
            
            # Line plot
            fig_line = go.Figure()
            for column in pivot_df.columns:
                fig_line.add_trace(go.Scatter(
                    x=pivot_df.index, 
                    y=pivot_df[column],
                    mode='lines+markers',
                    name=f'Thickness {column} mm'
                ))
            
            fig_line.update_layout(
                title='Rejection Percentage Trends by Thickness',
                xaxis_title='Date',
                yaxis_title='Rejection Percentage (%)',
                height=500,
                width=800
            )
            
            st.plotly_chart(fig_line)

if __name__ == '__main__':
    main()