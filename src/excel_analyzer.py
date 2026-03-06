import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
import warnings

warnings.filterwarnings('ignore')

# Configure Streamlit
st.set_page_config(
    page_title="Excel Analyzer - Demand Consensus",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .metric-card {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 8px;
        margin: 10px 0;
    }
    .outlier-high {
        background-color: #ffcccc;
        padding: 10px;
        border-radius: 5px;
        border-left: 5px solid #ff3333;
    }
    .outlier-low {
        background-color: #ccffcc;
        padding: 10px;
        border-radius: 5px;
        border-left: 5px solid #33ff33;
    }
    </style>
""", unsafe_allow_html=True)

def load_excel_file(uploaded_file):
    """Load Excel file and return sheet names and data"""
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        return excel_file.sheet_names, excel_file
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None, None

def compare_excel_files(df1, df2, threshold):
    """Compare two dataframes and identify outliers"""
    # Ensure both dataframes have the same structure
    if 'Key' not in df1.columns or 'Key' not in df2.columns:
        st.error("Both Excel files must have a 'Key' column (Column A)")
        return None
    
    # Set 'Key' as index for comparison
    df1_indexed = df1.set_index('Key')
    df2_indexed = df2.set_index('Key')
    
    # Find common keys
    common_keys = set(df1_indexed.index) & set(df2_indexed.index)
    
    if len(common_keys) == 0:
        st.error("No matching Keys found between the two Excel files")
        return None
    
    # Create comparison dataframe
    comparison_data = []
    
    for key in sorted(common_keys):
        row_data = {'Key': key}
        
        # Get numeric columns
        numeric_cols1 = df1_indexed.select_dtypes(include=[np.number]).columns
        numeric_cols2 = df2_indexed.select_dtypes(include=[np.number]).columns
        common_numeric_cols = set(numeric_cols1) & set(numeric_cols2)
        
        for col in common_numeric_cols:
            try:
                val1 = float(df1_indexed.loc[key, col])
                val2 = float(df2_indexed.loc[key, col])
                
                if val1 != 0:
                    pct_diff = abs((val2 - val1) / val1) * 100
                else:
                    pct_diff = 0 if val2 == 0 else 100
                
                row_data[f'{col}_File1'] = val1
                row_data[f'{col}_File2'] = val2
                row_data[f'{col}_Diff%'] = pct_diff
                row_data[f'{col}_IsOutlier'] = pct_diff > threshold
                
            except (ValueError, TypeError):
                continue
        
        if len(row_data) > 1:  # Only add if there's data
            comparison_data.append(row_data)
    
    return pd.DataFrame(comparison_data)

def get_plant_summary(df, plant_column=None):
    """Generate plant-based summary"""
    if plant_column is None or plant_column not in df.columns:
        return None
    
    summary_data = []
    for plant in df[plant_column].unique():
        if pd.isna(plant):
            continue
        
        plant_df = df[df[plant_column] == plant]
        
        # Count outliers
        outlier_cols = [col for col in plant_df.columns if 'IsOutlier' in col]
        if outlier_cols:
            outlier_count = plant_df[outlier_cols].sum().sum()
            total_comparisons = len(plant_df) * len(outlier_cols)
            outlier_pct = (outlier_count / total_comparisons * 100) if total_comparisons > 0 else 0
        else:
            outlier_count = 0
            outlier_pct = 0
        
        summary_data.append({
            'Plant': plant,
            'Records': len(plant_df),
            'Outliers': int(outlier_count),
            'Outlier %': round(outlier_pct, 2)
        })
    
    return pd.DataFrame(summary_data)

def highlight_outliers(df):
    """Highlight outlier rows"""
    outlier_cols = [col for col in df.columns if 'IsOutlier' in col]
    if not outlier_cols:
        return df
    
    def row_highlighter(row):
        if any(row[col] for col in outlier_cols if col in row.index):
            return ['background-color: #ffe6e6'] * len(row)
        return [''] * len(row)
    
    return df.style.apply(row_highlighter, axis=1)

# Main App
st.title("📊 Excel Analyzer - Demand Consensus")
st.markdown("Compare two Excel files and identify outliers by percentage threshold")

# Sidebar Configuration
st.sidebar.header("⚙️ Configuration")

# File Upload Section
st.sidebar.subheader("📁 File Upload")
col1, col2 = st.sidebar.columns(2)

with col1:
    file1 = st.file_uploader("Upload File 1", type=['xlsx', 'xls', 'csv'], key='file1')

with col2:
    file2 = st.file_uploader("Upload File 2", type=['xlsx', 'xls', 'csv'], key='file2')

if file1 and file2:
    # Load files
    sheets1, excel1 = load_excel_file(file1)
    sheets2, excel2 = load_excel_file(file2)
    
    if sheets1 and sheets2:
        # Sheet Selection
        st.sidebar.subheader("📄 Sheet Selection")
        col1, col2 = st.sidebar.columns(2)
        
        with col1:
            sheet1_name = st.selectbox("Select Sheet from File 1", sheets1, key='sheet1')
        
        with col2:
            sheet2_name = st.selectbox("Select Sheet from File 2", sheets2, key='sheet2')
        
        # Load selected sheets
        df1 = pd.read_excel(excel1, sheet_name=sheet1_name)
        df2 = pd.read_excel(excel2, sheet_name=sheet2_name)
        
        # Threshold Selection
        st.sidebar.subheader("🎯 Outlier Threshold")
        threshold_option = st.sidebar.radio(
            "Select threshold percentage:",
            [">5%", ">10%", "Custom"]
        )
        
        if threshold_option == ">5%":
            threshold = 5
        elif threshold_option == ">10%":
            threshold = 10
        else:
            threshold = st.sidebar.slider("Set custom threshold (%)", 1, 50, 5)
        
        st.sidebar.info(f"📌 Current Threshold: {threshold}%")
        
        # Plant column selection
        st.sidebar.subheader("🏭 Plant Column")
        plant_col_options = [None] + df1.columns.tolist()
        plant_column = st.sidebar.selectbox(
            "Select Plant/Location column for grouping",
            plant_col_options,
            format_func=lambda x: "None (No grouping)" if x is None else x
        )
        
        # Perform comparison
        comparison_df = compare_excel_files(df1, df2, threshold)
        
        if comparison_df is not None:
            st.success(f"✅ Comparison complete - {len(comparison_df)} records with matching Keys found")
            
            # Tabs for different views
            tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview", "🎯 Outliers", "🏭 Plant Summary", "📋 Detailed View"])
            
            with tab1:
                st.subheader("Comparison Overview")
                
                # Get numeric columns for statistics
                numeric_cols = [col for col in comparison_df.columns if col.endswith('_Diff%')]
                
                if numeric_cols:
                    col1, col2, col3, col4 = st.columns(4)
                    
                    all_diffs = comparison_df[numeric_cols].values.flatten()
                    all_diffs = all_diffs[~np.isnan(all_diffs)]
                    
                    with col1:
                        st.metric("Total Records", len(comparison_df))
                    
                    with col2:
                        outlier_count = comparison_df[[col for col in comparison_df.columns if col.endswith('_IsOutlier')]].sum().sum()
                        st.metric("Total Outliers", int(outlier_count))
                    
                    with col3:
                        st.metric("Avg Difference %", f"{all_diffs.mean():.2f}%")
                    
                    with col4:
                        st.metric("Max Difference %", f"{all_diffs.max():.2f}%")
                    
                    # Distribution chart
                    st.subheader("Distribution of Differences")
                    fig = go.Figure(data=[
                        go.Histogram(x=all_diffs, nbinsx=30, name="Difference %")
                    ])
                    fig.update_layout(
                        xaxis_title="Percentage Difference (%)",
                        yaxis_title="Count",
                        hovermode="x unified",
                        height=400
                    )
                    st.plotly_chart(fig, use_container_width=True)
            
            with tab2:
                st.subheader("🎯 Outlier Analysis")
                
                # Find outlier columns
                outlier_cols = [col for col in comparison_df.columns if col.endswith('_IsOutlier')]
                
                if outlier_cols:
                    # Get records with outliers
                    outlier_mask = comparison_df[outlier_cols].any(axis=1)
                    outlier_records = comparison_df[outlier_mask]
                    
                    st.info(f"Found {len(outlier_records)} records with outliers (threshold: {threshold}%)")
                    
                    # Create outlier summary by column
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Count outliers by metric
                        outlier_counts = {}
                        for col in outlier_cols:
                            metric_name = col.replace('_IsOutlier', '')
                            outlier_counts[metric_name] = comparison_df[col].sum()
                        
                        if outlier_counts:
                            fig = go.Figure(data=[
                                go.Bar(x=list(outlier_counts.keys()), y=list(outlier_counts.values()))
                            ])
                            fig.update_layout(
                                title="Outliers by Metric",
                                xaxis_title="Metric",
                                yaxis_title="Count",
                                height=400
                            )
                            st.plotly_chart(fig, use_container_width=True)
                    
                    with col2:
                        # Top outlier differences
                        diff_cols = [col for col in comparison_df.columns if col.endswith('_Diff%')]
                        if diff_cols:
                            all_diffs_with_idx = []
                            for idx, row in comparison_df.iterrows():
                                for col in diff_cols:
                                    if row[[c for c in comparison_df.columns if c.endswith('_IsOutlier') and c.replace('_IsOutlier', '') == col.replace('_Diff%', '')][0]] if [c for c in comparison_df.columns if c.endswith('_IsOutlier') and c.replace('_IsOutlier', '') == col.replace('_Diff%', '')] else False:
                                        all_diffs_with_idx.append({'Key': row['Key'], 'Metric': col.replace('_Diff%', ''), 'Diff%': row[col]})
                            
                            if all_diffs_with_idx:
                                top_outliers = sorted(all_diffs_with_idx, key=lambda x: x['Diff%'], reverse=True)[:10]
                                top_df = pd.DataFrame(top_outliers)
                                
                                fig = go.Figure(data=[
                                    go.Bar(x=top_df['Key'] + ' - ' + top_df['Metric'], y=top_df['Diff%'])
                                ])
                                fig.update_layout(
                                    title="Top 10 Outlier Differences",
                                    xaxis_title="Key - Metric",
                                    yaxis_title="Difference %",
                                    height=400
                                )
                                st.plotly_chart(fig, use_container_width=True)
                    
                    st.subheader("Outlier Records")
                    
                    # Display outlier records with highlighting
                    display_cols = ['Key'] + [col for col in comparison_df.columns if not col.endswith('_IsOutlier')]
                    st.dataframe(
                        comparison_df[outlier_mask][display_cols].style.highlight_max(axis=0),
                        use_container_width=True
                    )
                else:
                    st.info("No outlier columns found")
            
            with tab3:
                if plant_column:
                    st.subheader(f"🏭 Summary by {plant_column}")
                    
                    plant_summary = get_plant_summary(comparison_df, plant_column)
                    
                    if plant_summary is not None and len(plant_summary) > 0:
                        # Add plant_column values from original df
                        plant_summary_with_data = comparison_df.groupby(plant_column).agg({
                            'Key': 'count'
                        }).rename(columns={'Key': 'Records'}).reset_index()
                        
                        plant_summary_with_data.columns = [plant_column, 'Records']
                        
                        # Count outliers per plant
                        outlier_cols = [col for col in comparison_df.columns if col.endswith('_IsOutlier')]
                        for plant in plant_summary_with_data[plant_column]:
                            plant_df = comparison_df[comparison_df[plant_column] == plant]
                            if outlier_cols:
                                outlier_count = plant_df[outlier_cols].sum().sum()
                            else:
                                outlier_count = 0
                            plant_summary_with_data.loc[plant_summary_with_data[plant_column] == plant, 'Outliers'] = int(outlier_count)
                        
                        plant_summary_with_data['Outlier %'] = (plant_summary_with_data['Outliers'] / plant_summary_with_data['Records'] * 100).round(2)
                        
                        # Display summary cards
                        cols = st.columns(len(plant_summary_with_data))
                        
                        for i, (idx, row) in enumerate(plant_summary_with_data.iterrows()):
                            with cols[i]:
                                st.metric(
                                    f"{row[plant_column]}",
                                    f"{row['Records']} records",
                                    f"{row['Outliers']} outliers ({row['Outlier %']:.1f}%)"
                                )
                        
                        # Bar chart comparison
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            fig = go.Figure(data=[
                                go.Bar(x=plant_summary_with_data[plant_column], y=plant_summary_with_data['Records'], name='Total Records')
                            ])
                            fig.update_layout(title=f"Records by {plant_column}", height=400)
                            st.plotly_chart(fig, use_container_width=True)
                        
                        with col2:
                            fig = go.Figure(data=[
                                go.Bar(x=plant_summary_with_data[plant_column], y=plant_summary_with_data['Outlier %'], name='Outlier %')
                            ])
                            fig.update_layout(title=f"Outlier % by {plant_column}", height=400)
                            st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("No plant summary data available")
                else:
                    st.info("Please select a Plant/Location column to see the summary")
            
            with tab4:
                st.subheader("📋 Detailed Comparison View")
                
                # Filter options
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    show_outliers_only = st.checkbox("Show outliers only", value=False)
                
                with col2:
                    if plant_column:
                        plant_filter = st.multiselect(
                            "Filter by Plant",
                            options=[None] + sorted(comparison_df[plant_column].dropna().unique().tolist())
                        )
                    else:
                        plant_filter = []
                
                with col3:
                    display_metrics = st.multiselect(
                        "Select metrics to display",
                        options=[col.replace('_Diff%', '') for col in comparison_df.columns if col.endswith('_Diff%')],
                        default=[col.replace('_Diff%', '') for col in comparison_df.columns if col.endswith('_Diff%')][:3]
                    )
                
                # Apply filters
                filtered_df = comparison_df.copy()
                
                if show_outliers_only:
                    outlier_cols = [col for col in filtered_df.columns if col.endswith('_IsOutlier')]
                    if outlier_cols:
                        filtered_df = filtered_df[filtered_df[outlier_cols].any(axis=1)]
                
                if plant_filter and plant_column:
                    # Remove None from filter
                    plant_filter = [p for p in plant_filter if p is not None]
                    if plant_filter:
                        filtered_df = filtered_df[filtered_df[plant_column].isin(plant_filter)]
                
                # Build display columns
                display_cols = ['Key']
                
                for metric in display_metrics:
                    display_cols.extend([
                        f'{metric}_File1',
                        f'{metric}_File2',
                        f'{metric}_Diff%'
                    ])
                
                if plant_column:
                    display_cols.insert(1, plant_column)
                
                display_cols = [col for col in display_cols if col in filtered_df.columns]
                
                st.dataframe(
                    filtered_df[display_cols].sort_values('Key'),
                    use_container_width=True,
                    height=600
                )
                
                # Download option
                csv = filtered_df[display_cols].to_csv(index=False)
                st.download_button(
                    label="📥 Download Comparison as CSV",
                    data=csv,
                    file_name="excel_comparison.csv",
                    mime="text/csv"
                )
else:
    st.info("👈 Please upload two Excel files using the sidebar to begin analysis")
    
    with st.expander("ℹ️ How to use this tool"):
        st.markdown("""
        1. **Upload Files**: Use the sidebar to upload two Excel files for comparison
        2. **Select Sheets**: Choose which sheet from each file you want to compare
        3. **Set Threshold**: Choose your outlier detection threshold (>5%, >10%, or custom)
        4. **Select Plant Column**: Choose the column that contains Plant/Location information
        5. **Analyze**: View results in different tabs:
           - **Overview**: General statistics and distribution
           - **Outliers**: Detailed outlier analysis
           - **Plant Summary**: Summary cards and charts by plant
           - **Detailed View**: Full detailed view with filtering options
        
        **Requirements:**
        - Both files must have a "Key" column (Column A)
        - Both files should have the same structure/columns
        - Numeric columns will be compared for percentage differences
        """)
