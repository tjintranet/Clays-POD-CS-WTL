# PoD Batch Reader - Excel to PDF Work List

A web-based application for processing Excel files containing Print-on-Demand (PoD) batch data and generating organized PDF work lists.

## Features

### Excel File Processing
- **File Support**: Reads `.xlsx` and `.xlsm` files with "Master list" sheet
- **Drag & Drop**: Easy file upload with drag-and-drop interface
- **Data Validation**: Automatically validates and cleans input data
- **Error Handling**: Clear error messages for missing sheets or invalid data

### Dual Data Processing
- **Main Batch Data**: Regular production batches (8pp, 12pp, etc.)
  - Uses "File Quant" (Column B) for quantity
  - Displays with DCLAY paper codes
  - Shows "Rows" in quantity column
- **4pp Reference Entries**: Special 4pp style entries
  - Uses "Bk Quant" (Column C) for quantity
  - Displays without paper codes
  - Shows "Print Qty" in quantity column
  - Highlighted with yellow warning styling

### Data Organization
- **Paper Type Grouping**: Groups entries by paper type (Bulky 52/115, Cream 65/138, etc.)
- **Smart Sorting**: 
  - Paper types sorted by DCLAY code order (01, 02, 03, 05)
  - Batches within each type sorted by text batch number (descending)
- **Summary Statistics**: Shows counts of paper types, batches, and total quantities

### PDF Generation
- **Comprehensive Output**: Includes both main batch data and 4pp entries
- **Section Separation**: 4pp entries forced to new page for clear distinction
- **Professional Formatting**: 
  - Headers and footers with page numbers
  - Section titles and summaries
  - Grouped by paper type with clear visual separation
- **Intelligent Filename**: Uses actual order date from Excel file
  - Format: `Clays_POD_WTL_YYYY-MM-DD.pdf`
  - Automatically parses various date formats from Excel

## Excel File Structure

The application expects an Excel file with a "Master list" sheet containing these columns:

| Column | Name | Description | Usage |
|--------|------|-------------|-------|
| A | Paper Type | Type of paper (e.g., "Bulky 52 / 115") | All entries |
| B | File Quant | File quantity | Main batch data |
| C | Bk Quant | Book quantity | 4pp entries only |
| D | Text | Text batch number | All entries |
| E | Covers | Cover batch number | All entries |
| F | Date | Order date | All entries |
| G | Style | Style indicator (8pp, 4pp, etc.) | Determines processing |

### Paper Type Mapping
- **Bulky 52 / 115** → DCLAY 01
- **Cream 65 / 138** → DCLAY 02
- **Book 52 / 82** → DCLAY 03
- **Book 55 / 108** → DCLAY 05

## How It Works

1. **Upload**: Drag and drop or select Excel file
2. **Processing**: Application automatically:
   - Reads the "Master list" sheet
   - Separates 4pp entries from main batch data
   - Uses appropriate quantity column for each type
   - Groups and sorts data by paper type and batch number
3. **Display**: Shows two sections:
   - Main batch data with DCLAY codes
   - 4pp entries (reference only) without codes
4. **Download**: Generate PDF with both sections, 4pp on separate page

## Technical Details

### Browser Compatibility
- Modern browsers with JavaScript enabled
- No server required - runs entirely in browser
- Uses HTML5 File API for file processing

### Libraries Used
- **Bootstrap 5.3.0**: UI framework and styling
- **jsPDF 2.5.1**: PDF generation
- **SheetJS (xlsx) 0.18.5**: Excel file parsing
- **Bootstrap Icons**: UI icons

### File Processing
- Client-side Excel parsing (no data sent to servers)
- Supports both .xlsx and .xlsm (macro-enabled) files
- Handles various date formats automatically
- Robust error handling for malformed data

## Usage Instructions

1. **Prepare Excel File**: Ensure your file has a "Master list" sheet with the required columns
2. **Upload**: Open the application and upload your Excel file
3. **Review**: Check the displayed data in both main and 4pp sections
4. **Download**: Click "Download PDF Work List" to generate the formatted PDF
5. **Clear**: Use "Clear All" to reset and process a new file

## Output Features

### Screen Display
- **Responsive Tables**: Scrollable tables with sticky headers
- **Visual Distinction**: 4pp entries highlighted in yellow
- **Summary Cards**: Statistics for each section
- **Paper Type Headers**: Clear grouping with batch counts

### PDF Output
- **Professional Layout**: Clean, printable format
- **Two Sections**: Main data and 4pp reference on separate pages
- **Complete Information**: All relevant batch details included
- **Page Management**: Automatic page breaks and continued headers
- **Date-based Naming**: Filename reflects actual order date

## Error Handling

The application handles various error conditions:
- Missing "Master list" sheet
- Invalid file formats
- Missing required data columns
- Invalid quantity values
- Corrupted Excel files
- Empty or malformed data

## Browser Requirements

- Modern browser (Chrome, Firefox, Safari, Edge)
- JavaScript enabled
- No additional plugins or software required
- Works offline after initial page load