// Paper type mapping and sort order
var paperTypeMapping = {
    'Bulky 52 / 115': { code: 'DCLAY 01', order: 1 },
    'Cream 65 / 138': { code: 'DCLAY 02', order: 2 },
    'Book 52 / 82': { code: 'DCLAY 03', order: 3 },
    'Book 55 / 108': { code: 'DCLAY 05', order: 4 }
};

function getPaperCode(paperType) {
    return paperTypeMapping[paperType] ? paperTypeMapping[paperType].code : 'Unknown';
}

function getPaperOrder(paperType) {
    return paperTypeMapping[paperType] ? paperTypeMapping[paperType].order : 999;
}

var batchData = [];
var fourppData = [];

document.addEventListener('DOMContentLoaded', function() {
    setupEventListeners();
});

function setupEventListeners() {
    var excelUploadArea = document.getElementById('excelUploadArea');
    var excelFileInput = document.getElementById('excelFileInput');
    
    excelUploadArea.addEventListener('click', function() { 
        excelFileInput.click(); 
    });
    excelUploadArea.addEventListener('dragover', handleDragOver);
    excelUploadArea.addEventListener('drop', handleExcelDrop);
    excelFileInput.addEventListener('change', handleExcelFileSelect);

    document.getElementById('downloadPdfBtn').addEventListener('click', downloadPDF);
    document.getElementById('clearAllBtn').addEventListener('click', clearAll);
}

function handleDragOver(e) {
    e.preventDefault();
    e.currentTarget.classList.add('dragover');
}

function handleExcelDrop(e) {
    e.preventDefault();
    e.currentTarget.classList.remove('dragover');
    var files = e.dataTransfer.files;
    if (files.length > 0) {
        handleExcelFile(files[0]);
    }
}

function handleExcelFileSelect(e) {
    if (e.target.files.length > 0) {
        handleExcelFile(e.target.files[0]);
    }
}

function handleExcelFile(file) {
    if (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xlsm')) {
        showStatus('Please select an Excel file (.xlsx or .xlsm).', 'danger');
        return;
    }

    showStatus('Processing Excel file...', 'info');

    var reader = new FileReader();
    reader.onload = function(e) {
        try {
            var data = new Uint8Array(e.target.result);
            
            // Enhanced workbook reading options for .xlsm files
            var workbook = XLSX.read(data, {
                type: 'array',
                cellStyles: true,
                cellFormulas: true,
                cellDates: true,
                cellNF: true,
                sheetStubs: true,
                bookVBA: true  // Support for macro-enabled files
            });
            
            console.log('Workbook loaded. Available sheets:', workbook.SheetNames);
            
            // Check if "Master list" sheet exists
            if (!workbook.Sheets['Master list']) {
                showStatus('Error: "Master list" sheet not found in the Excel file. Available sheets: ' + workbook.SheetNames.join(', '), 'danger');
                return;
            }

            var worksheet = workbook.Sheets['Master list'];
            console.log('Master list sheet found. Range:', worksheet['!ref']);
            
            // Convert with better options for Excel data
            var jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1,
                defval: '',
                blankrows: false,
                raw: false  // Format values as strings
            });
            
            console.log('Converted to JSON. Rows found:', jsonData.length);
            
            if (jsonData.length < 3) {
                showStatus('Excel file appears to be empty or missing data. Found ' + jsonData.length + ' rows.', 'danger');
                return;
            }

            processExcelData(jsonData);
            
        } catch (error) {
            console.error('Detailed Excel parsing error:', error);
            showStatus('Error reading Excel file: ' + error.message + '. Please ensure the file is not corrupted and contains a "Master list" sheet.', 'danger');
        }
    };
    
    reader.onerror = function() {
        showStatus('Error reading file. Please try again.', 'danger');
    };
    
    reader.readAsArrayBuffer(file);
}

function processExcelData(jsonData) {
    batchData = [];
    fourppData = [];
    
    console.log('Processing Excel data. Total rows:', jsonData.length);
    console.log('First few rows:', jsonData.slice(0, 5));
    
    // Skip row 1 (title) and row 2 (headers), start from row 3 (index 2)
    for (var i = 2; i < jsonData.length; i++) {
        var row = jsonData[i];
        
        console.log('Processing row', i + 1, ':', row);
        
        // Skip empty rows and total rows - but let's be more flexible
        if (!row || row.length === 0) {
            console.log('Skipping empty row', i + 1);
            continue;
        }
        
        // Check if it's a total row
        if (row[0] && row[0].toString().toLowerCase().includes('total')) {
            console.log('Skipping total row', i + 1, ':', row[0]);
            continue;
        }
        
        // Updated column mapping for new Excel structure:
        // A: Paper Type, B: File Quant, C: Bk Quant, D: Text, E: Covers, F: Date, G: Style
        var paperType = row[0] ? String(row[0]).trim() : '';
        var fileQuantity = row[1]; // File Quant (column B)
        var bkQuantity = row[2]; // Bk Quant (column C)
        var textBatch = row[3] ? String(row[3]).trim() : ''; // Text batch (column D)
        var coverBatch = row[4] ? String(row[4]).trim() : ''; // Cover batch (column E)
        var orderDate = row[5] ? String(row[5]).trim() : ''; // Date (column F)
        var style = row[6] ? String(row[6]).trim() : ''; // Style (column G)
        
        console.log('Row', i + 1, 'data:', {
            paperType: paperType,
            fileQuantity: fileQuantity,
            bkQuantity: bkQuantity,
            textBatch: textBatch,
            coverBatch: coverBatch,
            orderDate: orderDate,
            style: style
        });
        
        // Check if we have the minimum required data
        if (!paperType || !textBatch || !coverBatch) {
            console.log('Skipping row', i + 1, '- missing required data');
            continue;
        }
        
        // Determine which quantity to use based on style
        var quantity;
        var quantitySource;
        if (style && style.toLowerCase().trim() === '4pp') {
            quantity = bkQuantity; // Use Bk Quant for 4pp entries
            quantitySource = 'Bk Quant (Column C)';
        } else {
            quantity = fileQuantity; // Use File Quant for regular entries
            quantitySource = 'File Quant (Column B)';
        }
        
        // Check if quantity is valid
        if (!quantity) {
            console.log('Skipping row', i + 1, '- missing quantity in', quantitySource);
            continue;
        }
        
        // Check if quantity is a number (could be string from Excel)
        var numericQuantity = parseInt(quantity) || parseFloat(quantity) || 0;
        if (numericQuantity <= 0) {
            console.log('Skipping row', i + 1, '- invalid quantity:', quantity, 'from', quantitySource);
            continue;
        }
        
        var batchItem = {
            paperType: paperType,
            paperCode: getPaperCode(paperType),
            textBatchNumber: textBatch,
            coverBatchNumber: coverBatch,
            numberOfRows: numericQuantity,
            orderDate: orderDate,
            style: style,
            quantitySource: quantitySource // For debugging
        };
        
        // Separate 4pp entries from regular batch data
        // More specific check for 4pp - only exact match
        if (style && style.toLowerCase().trim() === '4pp') {
            console.log('Adding 4pp item (using Bk Quant):', batchItem);
            fourppData.push(batchItem);
        } else {
            console.log('Adding regular batch item (using File Quant):', batchItem);
            batchData.push(batchItem);
        }
    }

    console.log('Final processed batch data:', batchData);
    console.log('Total valid batches found:', batchData.length);
    console.log('Final processed 4pp data:', fourppData);
    console.log('Total 4pp entries found:', fourppData.length);
    
    displayResults();
}

function displayResults() {
    if (batchData.length === 0 && fourppData.length === 0) {
        showStatus('No valid batch data found in the Excel file.', 'warning');
        return;
    }

    // Process main batch data
    if (batchData.length > 0) {
        displayMainBatchData();
        document.getElementById('resultsCard').style.display = 'block';
    }

    // Process 4pp data
    if (fourppData.length > 0) {
        display4ppData();
        document.getElementById('fourppCard').style.display = 'block';
    }

    document.getElementById('downloadPdfBtn').disabled = false;
    
    // Clear upload status message
    document.getElementById('uploadStatus').innerHTML = '';
}

function displayMainBatchData() {
    // Group by paper type
    var groupedData = {};
    var totalRows = 0;

    batchData.forEach(function(item) {
        if (!groupedData[item.paperType]) {
            groupedData[item.paperType] = [];
        }
        groupedData[item.paperType].push(item);
        totalRows += item.numberOfRows;
    });

    // Sort paper types by specified order (DCLAY 01, 02, 03, 05)
    var paperTypes = Object.keys(groupedData).sort(function(a, b) {
        return getPaperOrder(a) - getPaperOrder(b);
    });
    
    // Sort batches within each paper type by text batch number in descending order
    paperTypes.forEach(function(paperType) {
        groupedData[paperType].sort(function(a, b) {
            var textBatchA = parseInt(a.textBatchNumber) || 0;
            var textBatchB = parseInt(b.textBatchNumber) || 0;
            return textBatchB - textBatchA; // Descending order
        });
    });

    displaySummary(paperTypes.length, batchData.length, totalRows);
    updateResultsDisplay(groupedData, paperTypes);
}

function display4ppData() {
    // Group 4pp data by paper type
    var groupedData = {};
    var totalRows = 0;

    fourppData.forEach(function(item) {
        if (!groupedData[item.paperType]) {
            groupedData[item.paperType] = [];
        }
        groupedData[item.paperType].push(item);
        totalRows += item.numberOfRows;
    });

    // Sort paper types by specified order (DCLAY 01, 02, 03, 05)
    var paperTypes = Object.keys(groupedData).sort(function(a, b) {
        return getPaperOrder(a) - getPaperOrder(b);
    });
    
    // Sort batches within each paper type by text batch number in descending order
    paperTypes.forEach(function(paperType) {
        groupedData[paperType].sort(function(a, b) {
            var textBatchA = parseInt(a.textBatchNumber) || 0;
            var textBatchB = parseInt(b.textBatchNumber) || 0;
            return textBatchB - textBatchA; // Descending order
        });
    });

    display4ppSummary(paperTypes.length, fourppData.length, totalRows);
    update4ppResultsDisplay(groupedData, paperTypes);
}

function updateResultsDisplay(groupedData, paperTypes) {
    var tbody = document.getElementById('resultsTableBody');
    tbody.innerHTML = '';

    paperTypes.forEach(function(paperType) {
        var items = groupedData[paperType];
        
        // Add paper type header row
        var headerRow = tbody.insertRow();
        headerRow.className = 'paper-type-header';
        var headerCell = headerRow.insertCell();
        headerCell.colSpan = 6;
        headerCell.innerHTML = '<strong>' + items[0].paperCode + ' - ' + paperType + ' (' + items.length + ' batches)</strong>';
        
        // Add data rows for this paper type
        items.forEach(function(item) {
            var row = tbody.insertRow();
            row.insertCell(0).textContent = item.paperCode;
            row.insertCell(1).textContent = item.paperType;
            row.insertCell(2).textContent = item.textBatchNumber;
            row.insertCell(3).textContent = item.coverBatchNumber;
            row.insertCell(4).textContent = item.numberOfRows;
            row.insertCell(5).textContent = item.orderDate;
        });
    });
}

function displaySummary(paperTypeCount, batchCount, totalRows) {
    var summary = document.getElementById('dataSummary');
    
    var summaryHtml = '<div class="row mb-3">';
    summaryHtml += '<div class="col-md-4">';
    summaryHtml += '<div class="alert alert-info">';
    summaryHtml += '<strong>Paper Types:</strong> ' + paperTypeCount;
    summaryHtml += '</div>';
    summaryHtml += '</div>';
    summaryHtml += '<div class="col-md-4">';
    summaryHtml += '<div class="alert alert-success">';
    summaryHtml += '<strong>Total Batches:</strong> ' + batchCount;
    summaryHtml += '</div>';
    summaryHtml += '</div>';
    summaryHtml += '<div class="col-md-4">';
    summaryHtml += '<div class="alert alert-warning">';
    summaryHtml += '<strong>Total Rows:</strong> ' + totalRows;
    summaryHtml += '</div>';
    summaryHtml += '</div>';
    summaryHtml += '</div>';

    summary.innerHTML = summaryHtml;
}

function display4ppSummary(paperTypeCount, batchCount, totalRows) {
    var summary = document.getElementById('fourppSummary');
    
    var summaryHtml = '<div class="alert alert-warning">';
    summaryHtml += '<strong>4pp Entries:</strong> ' + paperTypeCount + ' paper types, ';
    summaryHtml += batchCount + ' batches, ' + totalRows + ' total rows';
    summaryHtml += '</div>';

    summary.innerHTML = summaryHtml;
}

function update4ppResultsDisplay(groupedData, paperTypes) {
    var tbody = document.getElementById('fourppTableBody');
    tbody.innerHTML = '';

    paperTypes.forEach(function(paperType) {
        var items = groupedData[paperType];
        
        // Add paper type header row
        var headerRow = tbody.insertRow();
        headerRow.className = 'paper-type-header';
        var headerCell = headerRow.insertCell();
        headerCell.colSpan = 6;
        headerCell.innerHTML = '<strong>' + paperType + ' (' + items.length + ' batches - 4pp)</strong>';
        
        // Add data rows for this paper type
        items.forEach(function(item) {
            var row = tbody.insertRow();
            row.insertCell(0).textContent = ''; // Empty paper code for 4pp
            row.insertCell(1).textContent = item.paperType;
            row.insertCell(2).textContent = item.textBatchNumber;
            row.insertCell(3).textContent = item.coverBatchNumber;
            row.insertCell(4).textContent = item.numberOfRows;
            row.insertCell(5).textContent = item.orderDate;
        });
    });
}

function formatOrderDateForFilename(orderDateString) {
    if (!orderDateString) return new Date().toISOString().split('T')[0]; // fallback to today
    
    // Try to parse common date formats from Excel
    var dateStr = orderDateString.toLowerCase().trim();
    
    // Handle formats like "Tuesday 29th", "Wednesday 30th", "Friday 1st", etc.
    // Extract day number and determine month/year based on context
    var dayMatch = dateStr.match(/(\d{1,2})(st|nd|rd|th)/);
    if (dayMatch) {
        var day = parseInt(dayMatch[1]);
        var month, year = 2025; // Default year
        
        // Determine month based on day and current context
        if (dateStr.includes('july') || dateStr.includes('jul')) {
            month = 7;
        } else if (dateStr.includes('august') || dateStr.includes('aug')) {
            month = 8;
        } else if (dateStr.includes('september') || dateStr.includes('sep')) {
            month = 9;
        } else if (dateStr.includes('october') || dateStr.includes('oct')) {
            month = 10;
        } else {
            // Auto-detect based on day ranges (this is based on your example dates)
            if (day >= 29 && day <= 31) {
                month = 7; // July
            } else if (day >= 1 && day <= 31) {
                month = 8; // August (most common)
            } else {
                month = 8; // Default to August
            }
        }
        
        // Format as YYYY-MM-DD
        var formattedMonth = month.toString().padStart(2, '0');
        var formattedDay = day.toString().padStart(2, '0');
        return year + '-' + formattedMonth + '-' + formattedDay;
    }
    
    // Handle specific known dates from your examples
    if (dateStr.includes('29th')) {
        return '2025-07-29';
    } else if (dateStr.includes('30th')) {
        return '2025-07-30';
    } else if (dateStr.includes('31st')) {
        return '2025-07-31';
    } else if (dateStr.includes('1st')) {
        return '2025-08-01';
    } else if (dateStr.includes('2nd')) {
        return '2025-08-02';
    } else if (dateStr.includes('3rd')) {
        return '2025-08-03';
    } else if (dateStr.includes('12th')) {
        return '2025-08-12';
    }
    
    // Try to extract date if it's in a different format like DD/MM/YYYY
    var dateMatch = dateStr.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (dateMatch) {
        var day = dateMatch[1].padStart(2, '0');
        var month = dateMatch[2].padStart(2, '0');
        var year = dateMatch[3].length === 2 ? '20' + dateMatch[3] : dateMatch[3];
        return year + '-' + month + '-' + day;
    }
    
    // Try to extract just month and day if format is "Month DD"
    var monthDayMatch = dateStr.match(/(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+(\d{1,2})/);
    if (monthDayMatch) {
        var monthName = monthDayMatch[1];
        var day = parseInt(monthDayMatch[2]);
        var monthNum;
        
        switch(monthName) {
            case 'january': case 'jan': monthNum = 1; break;
            case 'february': case 'feb': monthNum = 2; break;
            case 'march': case 'mar': monthNum = 3; break;
            case 'april': case 'apr': monthNum = 4; break;
            case 'may': monthNum = 5; break;
            case 'june': case 'jun': monthNum = 6; break;
            case 'july': case 'jul': monthNum = 7; break;
            case 'august': case 'aug': monthNum = 8; break;
            case 'september': case 'sep': monthNum = 9; break;
            case 'october': case 'oct': monthNum = 10; break;
            case 'november': case 'nov': monthNum = 11; break;
            case 'december': case 'dec': monthNum = 12; break;
            default: monthNum = 8; // Default to August
        }
        
        var formattedMonth = monthNum.toString().padStart(2, '0');
        var formattedDay = day.toString().padStart(2, '0');
        return '2025-' + formattedMonth + '-' + formattedDay;
    }
    
    // Fallback to today's date
    console.log('Could not parse date string:', orderDateString, '- using today\'s date');
    return new Date().toISOString().split('T')[0];
}

function downloadPDF() {
    if (batchData.length === 0 && fourppData.length === 0) {
        alert('No data to download.');
        return;
    }

    var jsPDF = window.jspdf.jsPDF;
    var pdf = new jsPDF('p', 'mm', 'a4');

    console.log('Generating PDF with', batchData.length, 'main batches and', fourppData.length, '4pp batches');

    // PDF Header
    pdf.setFontSize(14);
    pdf.setFont(undefined, 'bold');
    pdf.text('PoD Work List - Batch Processing', 20, 20);

    // Get the order date from the Excel data for the header
    var orderDate = '';
    if (batchData.length > 0) {
        orderDate = batchData[0].orderDate;
    } else if (fourppData.length > 0) {
        orderDate = fourppData[0].orderDate;
    }
    
    if (orderDate) {
        pdf.setFontSize(10);
        pdf.setFont(undefined, 'normal');
        pdf.text('Order Date: ' + orderDate, 20, 30);
        var yPos = 45;
    } else {
        var yPos = 35;
    }

    // Column positions
    var colPositions = {
        paperCode: { x: 20 },
        paperType: { x: 50 },
        textBatch: { x: 100 },
        coverBatch: { x: 130 },
        rows: { x: 165 }
    };

    // Process main batch data first
    if (batchData.length > 0) {
        yPos = addBatchDataToPDF(pdf, batchData, 'Main Batch Data', yPos, colPositions);
    }

    // Process 4pp data - force onto new page
    if (fourppData.length > 0) {
        // Force new page for 4pp entries
        pdf.addPage();
        yPos = 20;
        yPos = addBatchDataToPDF(pdf, fourppData, '4pp for Titan', yPos, colPositions);
    }

    // Add page numbers
    var pageCount = pdf.internal.getNumberOfPages();
    for (var i = 1; i <= pageCount; i++) {
        pdf.setPage(i);
        pdf.setFontSize(8);
        pdf.text('Page ' + i + ' of ' + pageCount, 175, 287);
    }

    // Generate filename using order date from Excel
    var formattedDate = formatOrderDateForFilename(orderDate);
    var fileName = 'Clays_POD_WTL_' + formattedDate + '.pdf';
    
    console.log('Saving PDF as:', fileName);
    console.log('Order date from Excel:', orderDate);
    console.log('Formatted date for filename:', formattedDate);
    pdf.save(fileName);
}

function addBatchDataToPDF(pdf, dataArray, sectionTitle, startYPos, colPositions) {
    var yPos = startYPos;
    
    // Group data by paper type
    var groupedData = {};
    var totalRows = 0;
    var totalBatches = 0;

    dataArray.forEach(function(item) {
        if (!groupedData[item.paperType]) {
            groupedData[item.paperType] = [];
        }
        groupedData[item.paperType].push(item);
        totalRows += item.numberOfRows;
        totalBatches++;
    });

    // Sort paper types by specified order (DCLAY 01, 02, 03, 05)
    var paperTypes = Object.keys(groupedData).sort(function(a, b) {
        return getPaperOrder(a) - getPaperOrder(b);
    });
    
    // Sort batches within each paper type by text batch number in descending order
    paperTypes.forEach(function(paperType) {
        groupedData[paperType].sort(function(a, b) {
            var textBatchA = parseInt(a.textBatchNumber) || 0;
            var textBatchB = parseInt(b.textBatchNumber) || 0;
            return textBatchB - textBatchA;
        });
    });

    // Add section title
    if (yPos > 250) {
        pdf.addPage();
        yPos = 20;
    }
    
    pdf.setFontSize(12);
    pdf.setFont(undefined, 'bold');
    pdf.text(sectionTitle, 20, yPos);
    yPos += 10;

    // Add summary for this section
    pdf.setFontSize(10);
    pdf.setFont(undefined, 'normal');
    pdf.text('Paper Types: ' + paperTypes.length + ' | Batches: ' + totalBatches + ' | Rows: ' + totalRows, 20, yPos);
    yPos += 15;

    // Determine if this is 4pp section
    var is4ppSection = sectionTitle.includes('4pp');

    // Process each paper type
    paperTypes.forEach(function(paperType, groupIndex) {
        var items = groupedData[paperType];
        
        if (yPos > 250) {
            pdf.addPage();
            yPos = 20;
        }

        // Paper type header
        pdf.setFontSize(11);
        pdf.setFont(undefined, 'bold');
        pdf.setFillColor(220, 220, 220);
        pdf.rect(15, yPos - 3, 180, 10, 'F');
        pdf.setTextColor(0, 0, 0);
        
        // Different header format for 4pp vs main data
        var headerText;
        if (is4ppSection) {
            headerText = paperType + ' (' + items.length + ' batches - 4pp)';
        } else {
            headerText = items[0].paperCode + ' - ' + paperType + ' (' + items.length + ' batches)';
        }
        
        pdf.text(headerText, 17, yPos + 3);
        yPos += 15;

        // Column headers - different for 4pp vs main data
        pdf.setFontSize(8);
        pdf.setFont(undefined, 'bold');
        pdf.text('Paper Code', colPositions.paperCode.x, yPos);
        pdf.text('Paper Type', colPositions.paperType.x, yPos);
        pdf.text('Text Batch', colPositions.textBatch.x, yPos);
        pdf.text('Cover Batch', colPositions.coverBatch.x, yPos);
        
        // Different header for quantity column based on section
        if (is4ppSection) {
            pdf.text('Print Qty', colPositions.rows.x, yPos);
        } else {
            pdf.text('Rows', colPositions.rows.x, yPos);
        }

        pdf.setLineWidth(0.2);
        pdf.line(15, yPos + 2, 195, yPos + 2);
        yPos += 8;

        // Data rows
        pdf.setFont(undefined, 'normal');
        pdf.setFontSize(7);

        items.forEach(function(item) {
            if (yPos > 280) {
                pdf.addPage();
                yPos = 20;
                
                // Repeat header on new page
                pdf.setFontSize(11);
                pdf.setFont(undefined, 'bold');
                pdf.setFillColor(220, 220, 220);
                pdf.rect(15, yPos - 3, 180, 10, 'F');
                pdf.text(headerText + ' (continued)', 17, yPos + 3);
                yPos += 15;
                
                // Repeat column headers on new page
                pdf.setFontSize(8);
                pdf.text('Paper Code', colPositions.paperCode.x, yPos);
                pdf.text('Paper Type', colPositions.paperType.x, yPos);
                pdf.text('Text Batch', colPositions.textBatch.x, yPos);
                pdf.text('Cover Batch', colPositions.coverBatch.x, yPos);
                
                // Repeat appropriate quantity header
                if (is4ppSection) {
                    pdf.text('Print Qty', colPositions.rows.x, yPos);
                } else {
                    pdf.text('Rows', colPositions.rows.x, yPos);
                }
                
                pdf.line(15, yPos + 2, 195, yPos + 2);
                yPos += 8;
                pdf.setFontSize(7);
                pdf.setFont(undefined, 'normal');
            }

            var paperTypeText = item.paperType.length > 20 ? item.paperType.substring(0, 20) + '..' : item.paperType;

            // For 4pp entries, leave paper code blank and show paper type
            if (is4ppSection) {
                pdf.text('', colPositions.paperCode.x, yPos); // Empty paper code
                pdf.text(paperTypeText, colPositions.paperType.x, yPos);
            } else {
                pdf.text(item.paperCode, colPositions.paperCode.x, yPos);
                pdf.text(paperTypeText, colPositions.paperType.x, yPos);
            }
            
            pdf.text(item.textBatchNumber, colPositions.textBatch.x, yPos);
            pdf.text(item.coverBatchNumber, colPositions.coverBatch.x, yPos);
            pdf.text(String(item.numberOfRows), colPositions.rows.x, yPos);
            
            yPos += 5;
        });

        if (groupIndex < paperTypes.length - 1) {
            yPos += 10;
        }
    });

    return yPos;
}

function clearAll() {
    batchData = [];
    fourppData = [];
    
    document.getElementById('excelFileInput').value = '';
    document.getElementById('uploadStatus').innerHTML = '';
    document.getElementById('resultsCard').style.display = 'none';
    document.getElementById('fourppCard').style.display = 'none';
    document.getElementById('resultsTableBody').innerHTML = '';
    document.getElementById('fourppTableBody').innerHTML = '';
    document.getElementById('dataSummary').innerHTML = '';
    document.getElementById('fourppSummary').innerHTML = '';
    document.getElementById('downloadPdfBtn').disabled = true;
    
    showStatus('Application cleared successfully. Ready for new upload.', 'success');
    
    setTimeout(function() {
        document.getElementById('uploadStatus').innerHTML = '';
    }, 3000);
}

function showStatus(message, type) {
    var uploadStatus = document.getElementById('uploadStatus');
    var alertDiv = document.createElement('div');
    alertDiv.className = 'alert alert-' + type;
    alertDiv.textContent = message;
    uploadStatus.innerHTML = '';
    uploadStatus.appendChild(alertDiv);
}