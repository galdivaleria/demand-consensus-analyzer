# allow the module to be imported in environments where streamlit
# is not installed (e.g. during unit testing). If streamlit is missing we
# create a lightweight stub object that provides only the methods used below.
try:
    import streamlit as st
except ImportError:  # pragma: no cover - streamlit is a runtime dependency
    class _StreamlitStub:
        def __getattr__(self, name):
            # return self for chained attribute access (eg. st.sidebar.header)
            return self
        def __call__(self, *args, **kwargs):
            # calling any stubbed object is a no-op
            return None
    st = _StreamlitStub()

import pandas as pd
import numpy as np

# plotly is a runtime dependency for the Streamlit app but not needed in tests;
# provide simple stubs when it isn't installed.
try:
    import plotly.graph_objects as go
    import plotly.express as px
except ImportError:  # pragma: no cover - occurs only in test environments
    class _PlotlyStub:
        def __getattr__(self, name):
            def _noop(*args, **kwargs):
                return None
            return _noop
    go = _PlotlyStub()
    px = _PlotlyStub()

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
    /* Card styling for premium UI */
    .metric-card {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 8px;
        margin: 10px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }

    /* Outlier highlights */
    .outlier-high {
        background-color: #ffe6e6;
        padding: 10px;
        border-radius: 5px;
        border-left: 5px solid #ff3333;
    }
    .outlier-low {
        background-color: #e6ffe6;
        padding: 10px;
        border-radius: 5px;
        border-left: 5px solid #33ff33;
    }

    /* make dataframes scroll on small screens */
    div.stDataFrame > div {
        overflow-x: auto;
    }

    @media (max-width: 768px) {
        .stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4 {
            font-size: 1.2rem;
        }
    }
    </style>
""", unsafe_allow_html=True)


# the actual implementation is kept in a helper so we can apply caching only
# when streamlit is available (the stub sets cache_data to None which would
# otherwise break decorators during testing).
def _load_excel_file_impl(uploaded_file):
    """Load an Excel/CSV file and return sheet names along with the object used
    for subsequent reads.

    If a CSV is provided we return a single "sheet" and the dataframe itself.
    This keeps the API consistent with ``pd.ExcelFile`` which is returned for
    real Excel workbooks.
    """
    try:
        filename = getattr(uploaded_file, "name", "").lower()
        if filename.endswith('.csv'):
            # read csv directly
            df = pd.read_csv(uploaded_file)
            return ["Sheet1"], df
        # attempt Excel first; many bytes streams (like BytesIO in tests) may not
        # have a name, so ExcelFile will error on CSV; catch and try CSV fallback.
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            return excel_file.sheet_names, excel_file
        except Exception:
            # reset stream position and attempt CSV read
            try:
                if hasattr(uploaded_file, 'seek'):
                    uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file)
                return ["Sheet1"], df
            except Exception as e2:
                raise e2
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None, None

# caching was originally added for UI performance, but ExcelFile objects
# are not pickleable which causes failures in our test environment.  The
# complexity of working around that (e.g. serializing to bytes or avoiding
# ExcelFile altogether) isn't worth the trouble for such a small app, so we
# simply bypass caching.  Tests and users will always call the raw helper
# directly which ensures predictable behavior.
load_excel_file = _load_excel_file_impl

def filter_outliers_by_metric(df, metric_name=None):
    """Return boolean mask indicating rows considered outliers.

    If ``metric_name`` is provided the mask only reflects that metric's flag;
    otherwise any ``*_IsOutlier`` column will be used.  Returns a Series of
    booleans indexed the same as ``df``.
    """
    outlier_cols = [col for col in df.columns if col.endswith('_IsOutlier')]
    if metric_name:
        col = f"{metric_name}_IsOutlier"
        if col in df.columns:
            return df[col].astype(bool)
        else:
            return pd.Series(False, index=df.index)
    if outlier_cols:
        return df[outlier_cols].any(axis=1)
    return pd.Series(False, index=df.index)


def map_threshold_option(option, custom=None):
    """Convert a user-facing threshold option string into a numeric value.

    ``option`` is one of the radio button labels (e.g. ">5%", "Custom").
    ``custom`` is the value from the slider when the user selects "Custom".  A
    ``ValueError`` is raised for an unknown option or missing custom value.
    """
    # expanded set of thresholds to give users finer control
    mapping = {
        ">1%": 1,
        ">2%": 2,
        ">5%": 5,
        ">10%": 10,
        ">15%": 15,
        ">20%": 20,
        ">25%": 25,
        ">30%": 30,
        ">40%": 40,
        ">50%": 50,
        ">60%": 60,
        ">70%": 70,
        ">80%": 80,
        ">90%": 90,
    }
    if option in mapping:
        return mapping[option]
    if option == "Custom":
        if custom is None:
            raise ValueError("Custom threshold requires a numeric value")
        return custom
    raise ValueError(f"Unknown threshold option: {option}")

def compare_excel_files(df1, df2, threshold, plant_column=None, material_column=None):
    """Compare two dataframes and identify outliers.

    ``plant_column`` and ``material_column`` are optional column names that,
    when provided, will be carried through to the resulting comparison
    dataframe so that downstream summaries can be produced.  They must exist in
    ``df1`` (only the first file is considered for carrying metadata).
    """
    # validate threshold type
    if not isinstance(threshold, (int, float)):
        # allow numeric strings? not; enforce numeric
        raise TypeError("threshold must be a number")

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

        # propagate metadata if present
        if plant_column and plant_column in df1_indexed.columns:
            row_data[plant_column] = df1_indexed.loc[key, plant_column]
        if material_column and material_column in df1_indexed.columns:
            row_data[material_column] = df1_indexed.loc[key, material_column]
        
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

def get_plant_summary(df, plant_column=None, material_column=None):
    """Generate plant-based summary.

    If ``material_column`` is given and exists in ``df`` we also count the
    number of unique materials associated with each plant.
    """
    if plant_column is None or plant_column not in df.columns:
        return None
    
    include_materials = material_column and material_column in df.columns
    
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
        
        entry = {
            'Plant': plant,
            'Records': len(plant_df),
            'Outliers': int(outlier_count),
            'Outlier %': round(outlier_pct, 2)
        }
        if include_materials:
            entry['Materials'] = int(plant_df[material_column].nunique())
        summary_data.append(entry)
    
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


# Main App logic encapsulated in a function so that importing the module for tests
# doesn't execute Streamlit UI code. This makes the module safe to import in
# non-Streamlit environments and allows unit tests to exercise the core
# functionality without needing the full UI stack.

def run_app():
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
            
            # Load selected sheets or csvs
            if isinstance(excel1, pd.DataFrame):
                df1 = excel1
            else:
                df1 = pd.read_excel(excel1, sheet_name=sheet1_name)

            if isinstance(excel2, pd.DataFrame):
                df2 = excel2
            else:
                df2 = pd.read_excel(excel2, sheet_name=sheet2_name)
            
            # Threshold Selection
            st.sidebar.subheader("🎯 Outlier Threshold")
            threshold_option = st.sidebar.radio(
                "Select threshold percentage:",
                [
                    ">1%", ">2%", ">5%", ">10%", ">15%", ">20%", ">25%",
                    ">30%", ">40%", ">50%", ">60%", ">70%", ">80%", ">90%",
                    "Custom"
                ]
            )
            custom_val = None
            if threshold_option == "Custom":
                custom_val = st.sidebar.slider("Set custom threshold (%)", 1, 100, 5)
            threshold = map_threshold_option(threshold_option, custom_val)
            
            st.sidebar.info(f"📌 Current Threshold: {threshold}%")

            # allow users to download the comparison results once available
            if 'comparison_df' in locals() and comparison_df is not None:
                csv = comparison_df.to_csv(index=False).encode('utf-8')
                st.sidebar.download_button(
                    "Download Results as CSV",
                    data=csv,
                    file_name="comparison_results.csv",
                    mime="text/csv",
                )

            # download button for results after comparison
            if 'comparison_df' in locals() and comparison_df is not None:
                csv = comparison_df.to_csv(index=False).encode('utf-8')
                st.sidebar.download_button(
                    "Download Results as CSV",
                    data=csv,
                    file_name="comparison_results.csv",
                    mime="text/csv",
                )
            
            # Plant & material column selection
            st.sidebar.subheader("🏭 Plant Column")
            plant_col_options = [None] + df1.columns.tolist()
            plant_column = st.sidebar.selectbox(
                "Select Plant/Location column for grouping",
                plant_col_options,
                format_func=lambda x: "None (No grouping)" if x is None else x
            )

            st.sidebar.subheader("📦 Material Column")
            material_col_options = [None] + df1.columns.tolist()
            material_column = st.sidebar.selectbox(
                "Select Material column (for counting distinct values)",
                material_col_options,
                format_func=lambda x: "None" if x is None else x
            )
            
            # Perform comparison with spinner for better UX
            with st.spinner("🔍 Comparing files..."):
                comparison_df = compare_excel_files(df1, df2, threshold, plant_column, material_column)
            
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
                            height=400,
                            autosize=True
                        )
                        st.plotly_chart(fig, use_container_width=True)
                
                with tab2:
                    st.subheader("🎯 Outlier Analysis")
                    
                    # Find outlier columns
                    outlier_cols = [col for col in comparison_df.columns if col.endswith('_IsOutlier')]
                    metric_opts = [c.replace('_IsOutlier','') for c in outlier_cols]
                    selected_metric = st.selectbox(
                        "Show outliers for metric:",
                        options=["All"] + metric_opts,
                        index=0
                    )
                    
                    # determine which columns should be used for filtering
                    if selected_metric != "All":
                        outlier_filter_cols = [f"{selected_metric}_IsOutlier"]
                    else:
                        outlier_filter_cols = outlier_cols
                    
                    if outlier_filter_cols:
                        # Get records with outliers (possibly filtered)
                        outlier_mask = comparison_df[outlier_filter_cols].any(axis=1)
                        outlier_records = comparison_df[outlier_mask]
                        
                        st.info(f"Found {len(outlier_records)} records with outliers (threshold: {threshold}%)")
                        
                        # Create outlier summary by column
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            # Count outliers by metric (respect selected filter)
                            outlier_counts = {}
                            columns_to_count = outlier_filter_cols if selected_metric != "All" else outlier_cols
                            for col in columns_to_count:
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
                                    height=400,
                                    autosize=True,
                                )
                                st.plotly_chart(fig, use_container_width=True)
                        
                        with col2:
                            # Top outlier differences
                            diff_cols = [col for col in comparison_df.columns if col.endswith('_Diff%')]
                            if selected_metric != "All":
                                diff_cols = [f"{selected_metric}_Diff%"] if f"{selected_metric}_Diff%" in diff_cols else []
                            if diff_cols:
                                all_diffs_with_idx = []
                                for idx, row in comparison_df.iterrows():
                                    for col in diff_cols:
                                        is_outlier_col = f"{col.replace('_Diff%','')}_IsOutlier"
                                        if row.get(is_outlier_col, False):
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
                                        height=400,
                                        autosize=True,
                                    )
                                    st.plotly_chart(fig, use_container_width=True)
                        
                        st.subheader("Outlier Records")
                        
                        # Display outlier records with highlighting
                        display_cols = ['Key']
                        if plant_column:
                            display_cols.append(plant_column)
                        if material_column:
                            display_cols.append(material_column)
                        display_cols += [col for col in comparison_df.columns if not col.endswith('_IsOutlier')]
                        st.dataframe(
                            comparison_df[outlier_mask][display_cols].style.highlight_max(axis=0),
                            use_container_width=True
                        )
                    else:
                        st.info("No outlier columns found")
                
                with tab3:
                    if plant_column:
                        st.subheader(f"🏭 Summary by {plant_column}")
                        
                        plant_summary = get_plant_summary(comparison_df, plant_column, material_column)
                        
                        if plant_summary is not None and len(plant_summary) > 0:
                            # Add plant_column values from original df
                            plant_summary_with_data = comparison_df.groupby(plant_column).agg({
                                'Key': 'count'
                            }).rename(columns={'Key': 'Records'}).reset_index()
                            
                            plant_summary_with_data.columns = [plant_column, 'Records']
                            
                            # add distinct material counts if requested
                            if material_column and material_column in comparison_df.columns:
                                mats = comparison_df.groupby(plant_column)[material_column].nunique().reset_index()
                                mats.columns = [plant_column, 'Materials']
                                plant_summary_with_data = plant_summary_with_data.merge(mats, on=plant_column, how='left')
                            else:
                                plant_summary_with_data['Materials'] = 0
                            
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
                                    caption = f"{row['Outliers']} outliers ({row['Outlier %']:.1f}%)"
                                    if material_column:
                                        caption += f" · {int(row.get('Materials',0))} materials"
                                    st.metric(
                                        f"{row[plant_column]}",
                                        f"{row['Records']} records",
                                        caption
                                    )
                            
                            # Bar chart comparison
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                fig = go.Figure(data=[
                                    go.Bar(x=plant_summary_with_data[plant_column], y=plant_summary_with_data['Records'], name='Total Records')
                                ])
                                fig.update_layout(title=f"Records by {plant_column}", height=400, autosize=True)
                                st.plotly_chart(fig, use_container_width=True)
                            
                            with col2:
                                fig = go.Figure(data=[
                                    go.Bar(x=plant_summary_with_data[plant_column], y=plant_summary_with_data['Outlier %'], name='Outlier %')
                                ])
                                fig.update_layout(title=f"Outlier % by {plant_column}", height=400, autosize=True)
                                st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info("No plant summary data available")
                    else:
                        st.info("Please select a Plant/Location column to see the summary")
                
                with tab4:
                    st.subheader("📋 Detailed Comparison View")

                    metric_names = [col.replace('_Diff%', '') for col in comparison_df.columns if col.endswith('_Diff%')]
                    base_columns = ['Key']
                    if plant_column:
                        base_columns.append(plant_column)
                    if material_column:
                        base_columns.append(material_column)

                    # Selection controls
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        show_outliers_only = st.checkbox("Show outliers only", value=False)
                        selected_metrics = st.multiselect(
                            "Select metrics to analyze",
                            options=metric_names,
                            default=metric_names
                        )
                        group_by_column = st.selectbox(
                            "Group by column",
                            options=[None] + base_columns + metric_names,
                            format_func=lambda x: "None" if x is None else x
                        )

                    with col2:
                        selected_display_columns = st.multiselect(
                            "Select fields to display",
                            options=base_columns + [f"{m}_{suffix}" for m in metric_names for suffix in ["File1", "File2", "Diff%", "IsOutlier"]],
                            default=base_columns + [f"{m}_Diff%" for m in metric_names]
                        )

                    with col3:
                        selected_filter_columns = st.multiselect(
                            "Filter columns",
                            options=base_columns + metric_names,
                            default=[]
                        )
                        chart_metric = st.selectbox(
                            "Metric for chart",
                            options=selected_metrics if selected_metrics else metric_names
                        )
                        chart_type = st.selectbox(
                            "Chart type",
                            options=["Bar", "Line", "Scatter", "Histogram"]
                        )

                    # Apply filtering
                    filtered_df = comparison_df.copy()
                    if show_outliers_only:
                        outlier_cols = [c for c in filtered_df.columns if c.endswith('_IsOutlier')]
                        if outlier_cols:
                            filtered_df = filtered_df[outlier_cols].any(axis=1).pipe(lambda mask: filtered_df[mask])

                    for col in selected_filter_columns:
                        if col not in filtered_df.columns:
                            col = col if col in filtered_df.columns else col
                        if col not in filtered_df.columns:
                            continue

                        if pd.api.types.is_numeric_dtype(filtered_df[col]):
                            min_val = float(filtered_df[col].min(skipna=True)) if not filtered_df[col].isna().all() else 0.0
                            max_val = float(filtered_df[col].max(skipna=True)) if not filtered_df[col].isna().all() else 1.0
                            if min_val == max_val:
                                max_val = min_val + 1.0
                            selected_range = st.slider(
                                f"{col} range",
                                min_value=min_val,
                                max_value=max_val,
                                value=(min_val, max_val),
                                key=f"filter_{col}"
                            )
                            filtered_df = filtered_df[filtered_df[col].between(selected_range[0], selected_range[1])]
                        else:
                            options = sorted(filtered_df[col].dropna().unique().tolist())
                            chosen = st.multiselect(
                                f"{col} values",
                                options=options,
                                default=options,
                                key=f"filter_{col}_vals"
                            )
                            if chosen:
                                filtered_df = filtered_df[filtered_df[col].isin(chosen)]

                    # Determine display columns
                    if selected_display_columns:
                        cols_to_show = [c for c in selected_display_columns if c in filtered_df.columns]
                    else:
                        cols_to_show = ['Key'] + ([plant_column] if plant_column else []) + ([material_column] if material_column else [])
                        for m in selected_metrics:
                            cols_to_show += [f"{m}_File1", f"{m}_File2", f"{m}_Diff%", f"{m}_IsOutlier"]
                        cols_to_show = [c for c in cols_to_show if c in filtered_df.columns]

                    st.write(f"Showing {len(filtered_df)} records after filters")

                    if group_by_column:
                        if group_by_column in filtered_df.columns:
                            outlier_cols = [c for c in filtered_df.columns if c.endswith('_IsOutlier')]
                            grouped = filtered_df.groupby(group_by_column).agg(
                                Records=('Key', 'count'),
                                Outliers=(outlier_cols[0], 'sum') if outlier_cols else ('Key', 'count')
                            ).reset_index()
                            if outlier_cols:
                                grouped['Outliers'] = filtered_df.groupby(group_by_column)[outlier_cols].sum().sum(axis=1).values
                            grouped['Outlier %'] = (grouped['Outliers'] / grouped['Records'] * 100).round(2)

                            st.subheader(f"Grouped summary by {group_by_column}")
                            st.dataframe(grouped, use_container_width=True)

                            if chart_metric and chart_metric in metric_names:
                                chart_col = f"{chart_metric}_Diff%"
                                if chart_col in filtered_df.columns:
                                    fig_df = filtered_df.groupby(group_by_column)[chart_col].mean().reset_index()
                                    if chart_type == "Bar":
                                        fig = px.bar(fig_df, x=group_by_column, y=chart_col, title=f"Average {chart_metric} Diff% by {group_by_column}")
                                    elif chart_type == "Line":
                                        fig = px.line(fig_df, x=group_by_column, y=chart_col, title=f"Average {chart_metric} Diff% by {group_by_column}")
                                    elif chart_type == "Scatter":
                                        fig = px.scatter(fig_df, x=group_by_column, y=chart_col, title=f"Average {chart_metric} Diff% by {group_by_column}")
                                    else:
                                        fig = px.histogram(filtered_df, x=chart_col, title=f"Distribution of {chart_metric} Diff%")
                                    st.plotly_chart(fig, use_container_width=True)

                    if not group_by_column and chart_metric:
                        chart_col = f"{chart_metric}_Diff%"
                        if chart_col in filtered_df.columns:
                            if chart_type == "Bar":
                                fig = px.bar(filtered_df, x='Key', y=chart_col, title=f"{chart_metric} Diff% by Key")
                            elif chart_type == "Line":
                                fig = px.line(filtered_df, x='Key', y=chart_col, title=f"{chart_metric} Diff% by Key")
                            elif chart_type == "Scatter":
                                fig = px.scatter(filtered_df, x='Key', y=chart_col, title=f"{chart_metric} Diff% by Key")
                            else:
                                fig = px.histogram(filtered_df, x=chart_col, nbins=30, title=f"Distribution of {chart_metric} Diff%")
                            st.plotly_chart(fig, use_container_width=True)

                    st.dataframe(filtered_df[cols_to_show], use_container_width=True)

            # end of if comparison_df is not None
            else:
                st.error("Comparison could not be performed. Please check the input files.")

    else:
        st.info("Please upload both File 1 and File 2 to proceed.")
