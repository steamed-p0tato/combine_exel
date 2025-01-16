import streamlit as st
import pandas as pd
from io import BytesIO
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, AgGridTheme

st.set_page_config(
    page_title="Excel Combiner",
    layout="wide",
    initial_sidebar_state="collapsed"
)

def configure_grid(df):
    """Configure the interactive grid with simplified options"""
    gb = GridOptionsBuilder.from_dataframe(df)
    
    # Configure default column behavior
    gb.configure_default_column(
        editable=True,
        resizable=True,
        sorteable=True,
        wrapText=True,
        autoHeight=True
    )
    
    return gb.build()

def combine_excel_files(uploaded_files):
    """Combine multiple Excel files into a single DataFrame"""
    all_dfs = []
    
    with st.spinner("Processing files..."):
        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file)
                all_dfs.append(df)
            except Exception as e:
                st.error(f"Error reading {uploaded_file.name}: {str(e)}")
                return None
    
    if not all_dfs:
        return None
        
    combined_df = pd.concat(all_dfs, ignore_index=True)
    return combined_df

def main():
    # Custom CSS
    st.markdown("""
        <style>
        .main > div {
            padding-top: 2rem;
        }
        .stAlert {
            margin-top: 1rem;
            margin-bottom: 1rem;
        }
        .uploadedFile {
            margin-bottom: 0.5rem;
        }
        .ag-theme-streamlit {
            margin: 1rem 0;
        }
        </style>
    """, unsafe_allow_html=True)

    # Header
    st.title("📊 Excel File Combiner")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.write("Combine multiple Excel files into one and edit the results.")
        st.write("by Ariyan for Mam")
    with col2:
        # Theme selector
        theme = st.selectbox(
            "Select theme",
            ["STREAMLIT", "ALPINE", "BALHAM", "MATERIAL"],
            index=0,
            key="theme_selector"
        )

    # File upload section
    uploaded_files = st.file_uploader(
        "Drop your Excel files here",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="You can upload multiple Excel files at once"
    )

    if uploaded_files:
        # Show uploaded files in a clean format
        with st.expander("📁 Uploaded Files", expanded=True):
            for file in uploaded_files:
                st.markdown(f"<div class='uploadedFile'>📄 {file.name}</div>", unsafe_allow_html=True)
        
        if st.button("🔄 Combine Files", use_container_width=True):
            combined_df = combine_excel_files(uploaded_files)
            
            if combined_df is not None:
                st.markdown("### 📝 Preview and Edit")
                st.info("Double-click any cell to edit. Changes are saved automatically.")
                
                # Configure and display the interactive grid
                grid_options = configure_grid(combined_df)
                grid_response = AgGrid(
                    combined_df,
                    grid_options,
                    update_mode=GridUpdateMode.VALUE_CHANGED,
                    data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                    theme=getattr(AgGridTheme, theme),
                    height=500,
                    allow_unsafe_jscode=True,
                    reload_data=False
                )
                
                # Get the updated dataframe
                updated_df = pd.DataFrame(grid_response['data'])
                
                # Summary section
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Rows", len(updated_df))
                with col2:
                    st.metric("Total Columns", len(updated_df.columns))
                with col3:
                    st.metric("Files Combined", len(uploaded_files))
                
                # Create download button for combined file
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    updated_df.to_excel(writer, index=False, sheet_name='Combined Data')
                
                # Download section
                st.markdown("### 💾 Save Your Work")
                st.download_button(
                    label="Download Combined Excel File",
                    data=output.getvalue(),
                    file_name="combined_excel_files.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    else:
        # Show placeholder when no files are uploaded
        st.info("👆 Upload your Excel files to get started")

if __name__ == "__main__":
    main()
