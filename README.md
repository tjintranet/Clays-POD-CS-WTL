# PoD Batch Reader - Excel to PDF Work List Generator

A web application that reads Excel files containing Print on Demand (PoD) batch data and generates professional PDF work lists for production planning. The application processes batch information, organizes it by paper type with proper sorting, and creates downloadable work lists with correct filename formatting.

## Overview

This application allows production teams to quickly upload Excel batch files and generate organized PDF work lists. The system automatically reads batch data from the "Master list" sheet, groups batches by paper type (with assigned paper codes), sorts batch numbers in descending order, and creates professional PDF reports ready for printing and production use.

## Features

- **Excel File Processing** - Reads .xlsx and .xlsm files with "Master list" sheet
- **Intelligent Data Organization** - Groups batches by paper type with proper DCLAY codes
- **Smart Sorting** - Paper types in specified order, batch numbers in descending order within each group
- **Professional PDF Export** - Clean, printable work lists with order date in filename
- **Comprehensive Data Display** - Shows all essential batch information in browser and PDF
- **Responsive Design** - Works on desktop, tablet, and mobile devices
- **Error Handling** - Clear feedback for file format issues and missing data

## File Structure

```
project-folder/
├── index.html              # Main HTML interface
├── script.js               # JavaScript functionality  
└── README.md              # This documentation file
```

## Setup Instructions

1. **Download Files** - Save `index.html` and `script.js` to the same directory
2. **Open Application** - Open `index.html` in a web browser
3. **Upload Excel File** - Use drag & drop or click to browse for your Excel file

## Excel File Requirements

### File Format
- **File types**: .xlsx or .xlsm files
- **Required sheet**: Must contain a sheet named exactly "Master list"
- **Data structure**: Data should start from row 3 (row 1 = title, row 2 = headers)

### Expected Column Structure (Master list sheet)
| Column | Field Name | Description |
|--------|------------|-------------|
| A | Paper Type | e.g., "Bulky 52 / 115", "Book 55 / 108" |
| B | File Quant | Number of files/quantity to print |
| C | Text | Text batch number (6-digit number) |
| D | Covers | Cover batch number (6-digit number) |
| E | Date | Order date (e.g., "Tuesday 29th", "Friday 1st") |

### Example Excel Data
```
Row 1: TJ Continuous - PoD Work to Print - [Date]
Row 2: Paper Type | File Quant | Text | Covers | Date | ...
Row 3: Bulky 52 / 115 | 22 | 115841 | 115842 | Tuesday 29th | ...
Row 4: Bulky 52 / 115 | 26 | 115843 | 115844 | Tuesday 29th | ...
```

## Paper Type Mapping & Sort Order

The application maps paper types to DCLAY codes and displays them in this specific order:

1. **Bulky 52 / 115** → `DCLAY 01`
2. **Cream 65 / 138** → `DCLAY 02` 
3. **Book 52 / 82** → `DCLAY 03`
4. **Book 55 / 108** → `DCLAY 05`

Within each paper type group, batches are sorted by **Text Batch Number in descending order** (highest to lowest).

## Application Interface

### Browser Display
The web interface shows comprehensive batch information:
- **Paper Code** (DCLAY 01, 02, 03, 05)
- **Paper Type** (full description)
- **Seq** (sequence numbers 1, 2, 3... for reference)
- **Text Batch Number** 
- **Cover Batch Number**
- **No. of Rows** (quantity)
- **Order Date**

### PDF Export  
The generated PDF contains:
- **Paper Code** (DCLAY 01, 02, 03, 05)
- **Paper Type**
- **Text Batch Number** (bold)
- **Cover Batch Number** (bold)
- **Rows** (quantity)

*Note: Sequence numbers and Order Date are excluded from PDF to maintain clean, focused work lists.*

## Usage Instructions

### Step 1: Upload Excel File
1. **Drag and drop** your Excel file onto the upload area, or
2. **Click** the upload area to browse and select your file
3. **Wait** for processing confirmation message

### Step 2: Review Data
- **Check the summary** showing paper types, total batches, and total rows
- **Review the organized data** grouped by paper type codes
- **Verify batch numbers** are sorted correctly (descending within each group)
- **Confirm quantities and dates** are correctly imported

### Step 3: Generate PDF Work List
1. **Click "Download PDF Work List"** button
2. **PDF automatically downloads** with filename format: `Clays_POD_WTL_YYYY-MM-DD.pdf`
3. **Date in filename** comes from the Order Date in your Excel file

### Example Workflow
```
Upload: "PoD work Batches to print list 29th July.xlsm"
↓
Process: Reads "Master list" sheet, finds 32 batches across 4 paper types
↓
Display: Shows organized data with DCLAY codes and sequence numbers
↓
Export: Generates "Clays_POD_WTL_2025-07-29.pdf"
```

## PDF Output Features

### Professional Layout
- **A4 optimized** - Perfect for printing and production use
- **Clear headers** - Paper type sections with batch counts
- **Bold batch numbers** - Easy to read Text and Cover batch numbers
- **Organized sections** - Each paper type clearly separated
- **Page numbering** - Multi-page support with proper pagination

### Filename Convention
- **Format**: `Clays_POD_WTL_YYYY-MM-DD.pdf`
- **Date source**: Extracted from Excel Order Date field
- **Examples**: 
  - "Tuesday 29th" → `Clays_POD_WTL_2025-07-29.pdf`
  - "Friday 1st" → `Clays_POD_WTL_2025-08-01.pdf`

### Content Organization
1. **Header**: Work list title and order date
2. **Summary**: Total paper types, batches, and rows
3. **Grouped Sections**: 
   - DCLAY 01 - Bulky 52 / 115 (X batches)
   - DCLAY 02 - Cream 65 / 138 (X batches)
   - DCLAY 03 - Book 52 / 82 (X batches)  
   - DCLAY 05 - Book 55 / 108 (X batches)
4. **Data Rows**: Paper code, type, batch numbers, quantities

## Data Processing Logic

### File Reading
- **Automatic sheet detection** - Locates "Master list" sheet
- **Header skipping** - Ignores title row and header row
- **Data validation** - Checks for required fields (Paper Type, Quantity, Text, Covers)
- **Total row filtering** - Skips rows containing "Total" 

### Organization Rules
1. **Group by Paper Type** - All batches with same paper type together
2. **Apply Paper Codes** - Maps paper types to DCLAY codes  
3. **Sort Paper Types** - DCLAY 01, 02, 03, 05 order
4. **Sort Within Groups** - Text batch numbers descending (highest first)
5. **Generate Sequences** - Number 1, 2, 3... within each paper type group

### Quality Controls
- **Missing data validation** - Identifies incomplete rows
- **Type conversion** - Handles quantities as numbers or text
- **Error reporting** - Clear messages for data issues
- **Summary statistics** - Shows totals for verification

## Troubleshooting

### Excel File Issues

**"Master list sheet not found"**
- Ensure your Excel file contains a sheet named exactly "Master list"
- Check sheet tab name for extra spaces or different capitalization

**"No valid batch data found"**
- Verify data starts from row 3 (row 1 = title, row 2 = headers)
- Check that columns A, B, C, D contain required data
- Ensure quantities in column B are numbers, not text

**"Excel file appears to be empty"**
- File may be corrupted - try re-saving the Excel file
- Check that the file actually contains data beyond headers

### Data Display Issues

**Missing paper types in results**
- Unknown paper types will show as "Unknown" code
- Add new paper type mappings in script.js if needed

**Incorrect sorting**
- Batch numbers sort as text - ensure they're properly formatted numbers
- Check for leading/trailing spaces in batch number fields

**Wrong quantities**
- Verify column B contains numeric values
- Empty quantities will show as "0"

### PDF Generation Issues

**PDF won't download**
- Check browser popup blocker settings
- Ensure browser allows file downloads
- Try refreshing the page and re-processing

**Filename shows wrong date**
- Check Order Date format in Excel column E
- Supported formats: "Tuesday 29th", "Friday 1st", etc.
- Date parsing falls back to today's date if format not recognized

**PDF formatting issues**
- Large datasets may take time to generate
- Complex paper type names may be truncated for space
- Multi-page PDFs handle large batch lists automatically

## Browser Compatibility

### Supported Browsers
- **Chrome/Edge**: Full support with optimal performance
- **Firefox**: Full support 
- **Safari**: Full support
- **Mobile browsers**: Responsive design works on tablets and phones

### Requirements
- **JavaScript enabled** - Required for all functionality
- **Modern browser** - HTML5 file upload support needed
- **PDF viewer** - For viewing downloaded work lists

## Technical Details

### Technologies Used
- **HTML5** - Modern web standards with file upload
- **Bootstrap 5** - Responsive UI framework
- **JavaScript** - Client-side file processing
- **SheetJS (XLSX)** - Excel file reading and parsing
- **jsPDF** - PDF generation and download

### Performance Notes
- **Client-side processing** - All work done in browser for privacy
- **No server required** - Can run from any web server or locally
- **Memory efficient** - Optimized for large datasets
- **Fast processing** - Handles hundreds of batches quickly

### File Size Limits
- **Excel files**: Recommended under 10MB for optimal performance
- **Batch count**: Tested with 100+ batches across multiple paper types
- **PDF output**: Handles large datasets with automatic pagination

## Security & Privacy

- **No server uploads** - Files processed entirely in your browser
- **No data storage** - Information is not saved or transmitted
- **Local processing** - Complete privacy for sensitive batch data
- **No internet required** - Works offline after initial page load

## Customization

### Adding New Paper Types
To add new paper type mappings, edit the `paperTypeMapping` object in `script.js`:

```javascript
var paperTypeMapping = {
    'Bulky 52 / 115': { code: 'DCLAY 01', order: 1 },
    'Cream 65 / 138': { code: 'DCLAY 02', order: 2 },
    'Book 52 / 82': { code: 'DCLAY 03', order: 3 },
    'Book 55 / 108': { code: 'DCLAY 05', order: 4 },
    'New Paper Type': { code: 'DCLAY 06', order: 5 }  // Add new types here
};
```

### Modifying Date Formats
To support new date formats, update the `formatOrderDateForFilename` function in `script.js`:

```javascript
// Add new date format recognition
if (dateStr.includes('4th')) {
    return '2025-08-04';
}
```

## Version History

### Version 1.0 - Initial Release
- Excel file reading and processing
- Paper type organization and sorting
- PDF generation with proper formatting
- Responsive web interface

### Version 1.1 - Enhanced Organization  
- Added paper code mapping (DCLAY 01, 02, 03, 05)
- Implemented custom paper type sort order
- Added sequence numbers for reference
- Improved PDF layout and formatting

### Version 1.2 - Production Ready
- Optimized PDF generation for production use
- Added order date extraction for filenames
- Simplified PDF layout for clarity
- Enhanced error handling and user feedback

## Support

For technical issues or feature requests:
1. **Check this documentation** for common solutions
2. **Verify Excel file format** matches requirements
3. **Test with sample data** to isolate issues
4. **Contact system administrator** for installation help

## License

This application is provided for internal production use. Modify as needed for your specific workflow requirements.

---

**PoD Batch Reader** - Streamlining batch processing from Excel to production-ready work lists.