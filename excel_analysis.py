import streamlit as st
import pandas as pd
import io
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

# --- Page Config ---
st.set_page_config(page_title="Excel Filter & Export Tool", layout="wide")

# --- Helper Functions ---

@st.cache_data
def load_data(file, header_row=0, sheet_name=0):
    """Loads the Excel file into a dataframe."""
    try:
        return pd.read_excel(file, header=header_row, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

def convert_df_to_excel(df):
    """Converts dataframe to Excel bytes for download."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Filtered_Data')
    return output.getvalue()

def convert_df_to_pdf(df):
    """Converts dataframe to PDF bytes for download using ReportLab."""
    output = io.BytesIO()
    
    # Use landscape to fit more columns
    doc = SimpleDocTemplate(output, pagesize=landscape(letter))
    elements = []
    
    # Add Title
    styles = getSampleStyleSheet()
    title = Paragraph("Filtered Data Report", styles['Title'])
    elements.append(title)
    
    # Prepare data for Table (Header + Rows)
    # Convert all data to string to ensure compatibility with ReportLab
    data = [df.columns.to_list()] + df.astype(str).values.tolist()
    
    # Create Table
    # Layout calculation: simplistic approach, might wrap on very wide tables
    t = Table(data)
    
    # Add style
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 8), # Smaller font for data
    ])
    t.setStyle(style)
    
    elements.append(t)
    
    try:
        doc.build(elements)
    except Exception as e:
        return None 
        
    return output.getvalue()

# --- Main Interface ---

st.title("📊 Excel Data Filter & Export Studio")
st.markdown("Upload your Excel sheet, filter rows/columns, and export the result.")

# 1. File Upload Section
st.header("1. Upload File")
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # 1. Get Sheet Names
    try:
        xl_file = pd.ExcelFile(uploaded_file)
        sheet_names = xl_file.sheet_names
    except Exception as e:
        st.error(f"Error reading Excel file structure: {e}")
        st.stop()

    # 2. Settings Columns (Sheet & Header)
    col_settings_1, col_settings_2 = st.columns(2)
    
    with col_settings_1:
        selected_sheet = st.selectbox(
            "Select Worksheet",
            options=sheet_names
        )

    with col_settings_2:
        header_row_index = st.number_input(
            "Select Header Row Index (0 for first row, etc.)",
            min_value=0,
            value=0,
            step=1,
            help="Change this if your column names are not in the first row."
        )

    # 3. Load Data
    # Reset file pointer to beginning after reading sheet names
    uploaded_file.seek(0)
    
    df_original = load_data(uploaded_file, header_row=header_row_index, sheet_name=selected_sheet)
    
    if df_original is not None:
        st.success("File uploaded successfully!")
        
        # Initialize the filtered dataframe
        df_filtered = df_original.copy()

        # Layout: Split into sidebar (Filters) and Main Area (Display)
        
        # --- Sidebar: Row Filtering ---
        st.sidebar.header("Filter Rows")
        st.sidebar.markdown("Select columns to filter by value:")
        
        # Step 1: Choose which columns to apply filters to
        # We assume categorical filtering for simplicity in this UI
        filter_cols = st.sidebar.multiselect(
            "Choose columns to filter rows:",
            options=df_filtered.columns
        )
        
        # Step 2: Generate dynamic widgets for selected columns
        for col in filter_cols:
            unique_values = df_original[col].unique()
            selected_values = st.sidebar.multiselect(
                f"Select values for '{col}'",
                options=unique_values,
                default=unique_values
            )
            # Apply Filter
            df_filtered = df_filtered[df_filtered[col].isin(selected_values)]

        # --- Main Area: Column Selection & Preview ---
        
        st.header("2. Select Columns")
        all_columns = df_original.columns.tolist()
        selected_columns = st.multiselect(
            "Choose which columns to include in the final view:",
            options=all_columns,
            default=all_columns
        )
        
        # Apply Column Selection
        if selected_columns:
            df_final = df_filtered[selected_columns]
        else:
            st.warning("Please select at least one column.")
            df_final = pd.DataFrame()

        # --- Display Data ---
        st.header("3. Data Preview")
        st.write(f"Showing {len(df_final)} rows and {len(df_final.columns)} columns.")
        st.dataframe(df_final, use_container_width=True)

        # --- Export Section ---
        st.header("4. Export Data")
        
        if not df_final.empty:
            col1, col2 = st.columns(2)
            
            # Excel Export
            with col1:
                excel_data = convert_df_to_excel(df_final)
                st.download_button(
                    label="📥 Download as Excel (.xlsx)",
                    data=excel_data,
                    file_name="filtered_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # PDF Export
            with col2:
                # PDF generation can be tricky with too many columns
                if len(df_final.columns) > 10:
                    st.warning("⚠️ Warning: PDF export may look cluttered with more than 10 columns.")
                
                if st.button("Generate PDF Preview"):
                    pdf_data = convert_df_to_pdf(df_final)
                    if pdf_data:
                        st.download_button(
                            label="📥 Download as PDF",
                            data=pdf_data,
                            file_name="filtered_data.pdf",
                            mime="application/pdf"
                        )
                    else:
                        st.error("Could not generate PDF. The table might be too wide or contain incompatible characters.")
        else:
            st.info("No data to export based on current filters.")

else:
    st.info("Awaiting file upload...")