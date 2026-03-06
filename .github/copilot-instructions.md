# Excel Analyzer - Copilot Instructions

## Project Overview
Excel Analyzer is a Streamlit-based application for comparing two Excel files with advanced outlier detection and plant-based analysis. It allows users to identify discrepancies between datasets based on customizable percentage thresholds.

## Architecture & Key Features
- **File Upload**: Dual Excel file upload with validation
- **Sheet Selection**: Choose specific sheets from each workbook
- **Comparison Engine**: Numeric column comparison with percentage difference calculation
- **Outlier Detection**: Configurable thresholds (>5%, >10%, custom)
- **Plant Summary**: Aggregate metrics by plant/location
- **Visualizations**: Interactive charts using Plotly
- **Export**: CSV download functionality

## Tech Stack
- **Frontend**: Streamlit (Python web framework)
- **Data Processing**: Pandas, NumPy
- **Visualization**: Plotly
- **Excel Handling**: Openpyxl

## File Structure
```
src/excel_analyzer.py    - Main application
requirements.txt         - Dependencies
README.md               - User documentation
.github/copilot-instructions.md - This file
```

## Development Guidelines

### Adding New Features
1. Keep UI organized using Streamlit tabs and columns
2. Maintain responsive design for different screen sizes
3. Add new comparisons as separate functions
4. Document parameters and return types

### Error Handling
- Validate file uploads before processing
- Check for required "Key" column
- Provide user-friendly error messages
- Use st.error() for critical issues

### Performance Considerations
- Cache data loading with @st.cache_data for large files
- Filter before visualization to reduce rendering
- Use vectorized operations with pandas

## Common Tasks

### To Add a New Filter
1. Add selectbox/multiselect in the relevant tab
2. Apply filter before displaying data:
   ```python
   filtered_df = df[df[column].isin(values)]
   ```
3. Update display to show filtered results

### To Add a New Chart
1. Import go or px from plotly
2. Create figure with appropriate trace
3. Update layout for consistency
4. Use st.plotly_chart() to display

### To Add Export Format
1. Create format-specific conversion function
2. Add download button with appropriate mime type
3. Test file integrity after export

## Testing Checklist
- [ ] File upload validation
- [ ] Multiple sheet selection
- [ ] Outlier threshold changes
- [ ] Plant grouping with various plants
- [ ] CSV export functionality
- [ ] Chart rendering
- [ ] Filter application
- [ ] Different file formats (.xlsx, .xls, .csv)

## Future Enhancement Ideas
- Multiple sheet comparison in one operation
- Data quality reports
- Trend analysis over time
- Historical comparison archiving
- API integration for data sources
- Advanced statistical analysis options
- User preference saving
