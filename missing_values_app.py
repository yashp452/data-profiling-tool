import streamlit as st
import pandas as pd
import os
import tempfile
import zipfile
from io import BytesIO
import json
from ydata_profiling import ProfileReport

def analyze_missing_values_streamlit(excel_file, sheet_column_mapping):
    """
    Analyze missing values for Streamlit app with file downloads
    """
    # Create temporary directory for outputs
    temp_dir = tempfile.mkdtemp()
    summary_data = []
    html_reports = {}
    
    try:
        excel_file_obj = pd.ExcelFile(excel_file)
        
        for sheet_name, columns_to_analyze in sheet_column_mapping.items():
            if sheet_name not in excel_file_obj.sheet_names:
                st.warning(f"Sheet '{sheet_name}' not found in Excel file")
                continue
                
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                
                # Check if specified columns exist
                missing_cols = [col for col in columns_to_analyze if col not in df.columns]
                if missing_cols:
                    st.warning(f"Columns not found in sheet '{sheet_name}': {missing_cols}")
                    available_cols = [col for col in columns_to_analyze if col in df.columns]
                    if not available_cols:
                        continue
                    columns_to_analyze = available_cols
                
                df_selected = df[columns_to_analyze]
                
                if df_selected.empty:
                    continue
                
                total_rows = len(df_selected)
                
                # Collect summary data
                for col in columns_to_analyze:
                    missing_count = df_selected[col].isnull().sum()
                    missing_pct = (missing_count / total_rows) * 100
                    non_missing = total_rows - missing_count
                    
                    summary_data.append({
                        'Sheet': sheet_name,
                        'Column': col,
                        'Total_Rows': total_rows,
                        'Missing_Count': missing_count,
                        'Missing_Percentage': round(missing_pct, 2),
                        'Non_Missing_Count': non_missing
                    })
                
                # Generate HTML report
                try:
                    profile = ProfileReport(
                        df_selected,
                        title=f"Missing Values Analysis - {sheet_name}",
                        minimal=True,
                        explorative=False,
                    )
                    
                    html_content = profile.to_html()
                    html_reports[sheet_name] = html_content
                    
                except Exception as e:
                    st.error(f"Error generating HTML report for '{sheet_name}': {e}")
                
            except Exception as e:
                st.error(f"Error processing sheet '{sheet_name}': {e}")
                continue
        
        return summary_data, html_reports
        
    except Exception as e:
        st.error(f"Error analyzing file: {e}")
        return [], {}

def create_download_zip(summary_df, html_reports):
    """Create a zip file with all reports"""
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        # Add CSV summary
        csv_buffer = BytesIO()
        summary_df.to_csv(csv_buffer, index=False)
        zip_file.writestr('missing_values_summary.csv', csv_buffer.getvalue())
        
        # Add aggregated summary if multiple sheets
        if len(summary_df) > 0:
            agg_summary = summary_df.groupby('Column').agg({
                'Total_Rows': 'sum',
                'Missing_Count': 'sum',
                'Missing_Percentage': 'mean'
            }).round(2)
            
            agg_csv_buffer = BytesIO()
            agg_summary.to_csv(agg_csv_buffer)
            zip_file.writestr('aggregated_missing_summary.csv', agg_csv_buffer.getvalue())
        
        # Add HTML reports
        for sheet_name, html_content in html_reports.items():
            zip_file.writestr(f'{sheet_name}_missing_values_report.html', html_content)
    
    return zip_buffer.getvalue()

def main():
    st.set_page_config(
        page_title="Missing Values Analysis Tool",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Missing Values Analysis Tool")
    st.markdown("Upload an Excel file and analyze missing values across sheets and columns")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload an Excel file to analyze missing values"
    )
    
    if uploaded_file is not None:
        try:
            # Show file info
            st.success(f"✅ File uploaded: {uploaded_file.name}")
            
            # Load Excel file to show available sheets
            excel_file = pd.ExcelFile(uploaded_file)
            available_sheets = excel_file.sheet_names
            
            st.subheader("📋 Available Sheets")
            st.write(f"Found {len(available_sheets)} sheets: {', '.join(available_sheets)}")
            
            # Configuration section
            st.subheader("⚙️ Configuration")
            
            # Method to configure analysis
            config_method = st.radio(
                "Choose configuration method:",
                ["Interactive Setup", "JSON Configuration", "Quick Analysis (All Sheets, Same Columns)"]
            )
            
            sheet_column_mapping = {}
            
            if config_method == "Quick Analysis (All Sheets, Same Columns)":
                # Quick setup - same columns for all sheets
                st.markdown("**Quick Setup:** Analyze the same columns across all sheets")
                
                # Load first sheet to show available columns
                sample_df = pd.read_excel(uploaded_file, sheet_name=available_sheets[0])
                available_columns = list(sample_df.columns)
                
                selected_columns = st.multiselect(
                    f"Select columns to analyze (from {available_sheets[0]}):",
                    available_columns,
                    help="These columns will be analyzed in all sheets where they exist"
                )
                
                if selected_columns:
                    # Apply same columns to all sheets
                    for sheet in available_sheets:
                        sheet_column_mapping[sheet] = selected_columns
            
            elif config_method == "Interactive Setup":
                # Interactive setup - different columns for each sheet
                st.markdown("**Interactive Setup:** Choose specific columns for each sheet")
                
                for sheet in available_sheets:
                    with st.expander(f"📄 Configure Sheet: {sheet}"):
                        # Load sheet to show available columns
                        try:
                            sheet_df = pd.read_excel(uploaded_file, sheet_name=sheet)
                            sheet_columns = list(sheet_df.columns)
                            
                            selected_cols = st.multiselect(
                                f"Select columns for {sheet}:",
                                sheet_columns,
                                key=f"cols_{sheet}"
                            )
                            
                            if selected_cols:
                                sheet_column_mapping[sheet] = selected_cols
                                
                        except Exception as e:
                            st.error(f"Error loading sheet {sheet}: {e}")
            
            elif config_method == "JSON Configuration":
                # JSON configuration
                st.markdown("**JSON Configuration:** Paste or type your configuration")
                
                json_example = {
                    "Sheet1": ["Asset", "Super Reason"],
                    "Sheet2": ["Asset", "Super Reason", "Other Column"]
                }
                
                st.code(json.dumps(json_example, indent=2), language="json")
                
                json_input = st.text_area(
                    "Enter JSON configuration:",
                    height=150,
                    placeholder=json.dumps(json_example, indent=2)
                )
                
                if json_input:
                    try:
                        sheet_column_mapping = json.loads(json_input)
                        st.success("✅ JSON configuration loaded successfully")
                    except json.JSONDecodeError as e:
                        st.error(f"❌ Invalid JSON format: {e}")
            
            # Show current configuration
            if sheet_column_mapping:
                st.subheader("📋 Current Configuration")
                config_df = []
                for sheet, columns in sheet_column_mapping.items():
                    for col in columns:
                        config_df.append({"Sheet": sheet, "Column": col})
                
                if config_df:
                    st.dataframe(pd.DataFrame(config_df), use_container_width=True)
                
                # Analyze button
                if st.button("🚀 Run Missing Values Analysis", type="primary"):
                    with st.spinner("Analyzing missing values..."):
                        summary_data, html_reports = analyze_missing_values_streamlit(
                            uploaded_file, sheet_column_mapping
                        )
                        
                        if summary_data:
                            summary_df = pd.DataFrame(summary_data)
                            
                            # Display results
                            st.subheader("📊 Results")
                            
                            # Summary metrics
                            col1, col2, col3, col4 = st.columns(4)
                            
                            with col1:
                                st.metric("Sheets Processed", len(summary_df['Sheet'].unique()))
                            with col2:
                                st.metric("Columns Analyzed", len(summary_df['Column'].unique()))
                            with col3:
                                total_missing = summary_df['Missing_Count'].sum()
                                st.metric("Total Missing Values", f"{total_missing:,}")
                            with col4:
                                avg_missing_pct = summary_df['Missing_Percentage'].mean()
                                st.metric("Avg Missing %", f"{avg_missing_pct:.1f}%")
                            
                            # Detailed summary table
                            st.subheader("📋 Detailed Summary")
                            st.dataframe(summary_df, use_container_width=True)
                            
                            # Aggregated summary
                            if len(summary_df) > 1:
                                st.subheader("📈 Aggregated Summary by Column")
                                agg_summary = summary_df.groupby('Column').agg({
                                    'Total_Rows': 'sum',
                                    'Missing_Count': 'sum',
                                    'Missing_Percentage': 'mean'
                                }).round(2)
                                st.dataframe(agg_summary, use_container_width=True)
                            
                            # Charts
                            st.subheader("📊 Visualizations")
                            
                            # Missing percentage by column
                            if len(summary_df) > 0:
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    st.bar_chart(
                                        summary_df.set_index(['Sheet', 'Column'])['Missing_Percentage'],
                                        use_container_width=True
                                    )
                                    st.caption("Missing Percentage by Sheet and Column")
                                
                                with col2:
                                    # Missing count by sheet
                                    sheet_totals = summary_df.groupby('Sheet')['Missing_Count'].sum()
                                    st.bar_chart(sheet_totals, use_container_width=True)
                                    st.caption("Total Missing Values by Sheet")
                            
                            # Download section
                            st.subheader("💾 Download Reports")
                            
                            # Individual downloads
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                # CSV summary download
                                csv_data = summary_df.to_csv(index=False)
                                st.download_button(
                                    label="📄 Download Summary CSV",
                                    data=csv_data,
                                    file_name="missing_values_summary.csv",
                                    mime="text/csv"
                                )
                            
                            with col2:
                                # Aggregated summary download
                                if len(summary_df) > 1:
                                    agg_csv_data = agg_summary.to_csv()
                                    st.download_button(
                                        label="📈 Download Aggregated Summary",
                                        data=agg_csv_data,
                                        file_name="aggregated_missing_summary.csv",
                                        mime="text/csv"
                                    )
                            
                            # HTML reports downloads
                            if html_reports:
                                st.markdown("**Individual HTML Reports:**")
                                cols = st.columns(min(len(html_reports), 3))
                                for idx, (sheet_name, html_content) in enumerate(html_reports.items()):
                                    with cols[idx % 3]:
                                        st.download_button(
                                            label=f"📊 {sheet_name} Report",
                                            data=html_content,
                                            file_name=f"{sheet_name}_missing_values_report.html",
                                            mime="text/html"
                                        )
                            
                            # All reports as ZIP
                            st.markdown("**Download All Reports:**")
                            zip_data = create_download_zip(summary_df, html_reports)
                            st.download_button(
                                label="🗜️ Download All Reports (ZIP)",
                                data=zip_data,
                                file_name=f"missing_values_analysis_{uploaded_file.name.split('.')[0]}.zip",
                                mime="application/zip"
                            )
                            
                            st.success("🎉 Analysis completed successfully!")
                            
                        else:
                            st.error("❌ No data could be processed. Please check your configuration.")
            
            else:
                st.info("👆 Please configure which columns to analyze for each sheet above.")
                
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")
    
    else:
        st.info("👆 Please upload an Excel file to get started.")
        
        # Show example configuration
        st.subheader("💡 Example Usage")
        st.markdown("""
        1. **Upload** your Excel file
        2. **Configure** which columns to analyze for each sheet
        3. **Run analysis** to see missing values statistics  
        4. **Download** detailed reports (CSV summaries and HTML reports)
        
        **Configuration Options:**
        - **Quick Analysis**: Same columns across all sheets
        - **Interactive Setup**: Different columns for each sheet  
        - **JSON Configuration**: Advanced configuration with JSON
        """)

if __name__ == "__main__":
    main()