# Excel Analyzer - Demand Consensus

A powerful browser-based application for comparing two Excel files with advanced outlier detection and plant-based analysis. No installation required - just open the HTML file in your browser!

## Features

✨ **Core Functionality:**
- 📁 Upload and compare two Excel files
- 📄 Select specific sheets from each file
- 🎯 Customizable outlier detection (>5%, >10%, or custom threshold)
- 🏭 Plant/Location-based grouping and summary
- 📊 Interactive visualizations and charts
- 💾 Export results to CSV

🎯 **Outlier Detection:**
- Automatic comparison of numeric columns
- Percentage difference calculation based on ID (Key column)
- Configurable sensitivity thresholds
- Visual highlighting of outliers

📈 **Analysis Views:**
- **Overview Tab**: Statistics, distribution charts, and key metrics
- **Outliers Tab**: Detailed outlier analysis with top differences bar charts
- **Plant Summary**: Plant-based metrics cards and comparison charts
- **Detailed View**: Filterable detailed comparison with export options

## Quick Start

### No Installation Required!

1. **Download the project** to your local machine

2. **Open `index.html`** in your web browser (Chrome, Firefox, Edge, Safari)
   - Simply double-click the file, or
   - Right-click → Open with → Your preferred browser

3. **Start comparing Excel files!**

### That's it! No npm, Python, or dependencies needed.

## Usage

1. **Upload Excel files** using the file upload inputs in the sidebar

2. **Configure comparison**:
   - Select sheets from each file
   - Choose outlier threshold (>5%, >10%, or custom)
   - Select plant column for grouping (optional)

3. **Click "Compare Files"** button

4. **View results** in the different tabs:
   - 📊 **Overview**: General statistics and distribution
   - 🎯 **Outliers**: Detailed outlier findings
   - 🏭 **Plant Summary**: Plant-based aggregations and cards
   - 📋 **Detailed View**: Full comparison with filtering

5. **Download results** as CSV for further analysis

## File Requirements

### Column Structure
- **Column A (Key)**: Unique identifier for matching records between files
- **Other Columns**: Numeric and categorical data for comparison

### Supported Formats
- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)
- `.csv` (Comma-separated values)

### Example Structure
| Key | Plant | Metric1 | Metric2 | Metric3 |
|-----|-------|---------|---------|---------|
| ID001 | PlantA | 100 | 200 | 300 |
| ID002 | PlantB | 150 | 250 | 350 |

## Configuration

### Outlier Thresholds
- **Predefined**: >5%, >10%
- **Custom**: Set any threshold between 1% and 100%

### Plant Column
- Optional grouping parameter
- Enables summary cards and plant-based analysis
- Select "None" to skip plant-based aggregation

## Output

### Metrics Provided
- **Total Records**: Count of matched records
- **Outliers Detected**: Count of records exceeding threshold
- **Average Difference**: Mean percentage difference
- **Max Difference**: Highest percentage difference
- **Plant-based Outlier %**: Percentage of outliers per plant

### Export Options
- Download comparison results as CSV
- Filter by plant, metrics, or outlier status
- Customizable column selection

## Troubleshooting

### Common Issues

**"Error loading file"**
- Ensure file is a valid Excel or CSV format
- Check file is not corrupted
- Try in a different browser

**"Both Excel files must have a 'Key' column"**
- Column A must be named "Key" (case-sensitive)
- Check spelling and ensure it exists in both files

**"No matching Keys found"**
- The Key values don't match between files
- Verify data consistency between files
- Check for leading/trailing spaces in Key column

**Application doesn't load**
- Ensure you have a modern browser (Chrome, Firefox, Edge, Safari)
- Check internet connection (some CDN resources are loaded online)
- Try clearing browser cache

**Files don't upload**
- Check file size (most browsers support up to 4GB)
- Ensure file is not corrupted
- Try opening and resaving the Excel file

## Browser Compatibility

- ✅ Chrome/Chromium (recommended)
- ✅ Firefox
- ✅ Safari
- ✅ Edge
- ✅ Opera

**Recommended**: Chrome or Firefox for best performance

## Technologies Used (All from CDN - No Installation)

- **SheetJS**: Excel file reading and parsing
- **Chart.js**: Statistical visualizations
- **Plotly.js**: Interactive charts
- **Vanilla JavaScript**: Core application logic

## Project Structure

```
Analizator Demand Consensus/
├── index.html                  # Main application (open this!)
├── src/
│   └── app.js                  # Application logic
├── data/                       # Sample data directory
└── README.md                   # This file
```

## Performance Tips

- Works best with files up to 50MB
- For very large files, consider splitting by sheet
- Filtering by plant reduces computation time
- Chrome typically has the best performance

## Support

For issues or feature requests, check:
1. File format and structure (must be valid Excel or CSV)
2. Column naming (especially "Key" column - case-sensitive)
3. Browser compatibility
4. Internet connection for CDN resources

## Version History

- **v2.0.0** (2026): JavaScript/Browser version
  - No installation required
  - Runs entirely in the browser
  - All libraries via CDN
  
- **v1.0.0** (2024): Python/Streamlit version (archived)
  - Initial release with core features

## License

This project is created for Faurecia Demand Consensus Analysis.
