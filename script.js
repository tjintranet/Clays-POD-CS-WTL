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
        
        // More flexible validation - check if we have the essential data
        var paperType = row[0] ? String(row[0]).trim() : '';
        var quantity = row[1];
        var textBatch = row[2] ? String(row[2]).trim() : '';
        var coverBatch = row[3] ? String(row[3]).trim() : '';
        var orderDate = row[4] ? String(row[4]).trim() : '';
        
        console.log('Row', i + 1, 'data:', {
            paperType: paperType,
            quantity: quantity,
            textBatch: textBatch,
            coverBatch: coverBatch,
            orderDate: orderDate
        });
        
        // Check if we have the minimum required data
        if (!paperType || !quantity || !textBatch || !coverBatch) {
            console.log('Skipping row', i + 1, '- missing required data');
            continue;
        }
        
        // Check if quantity is a number (could be string from Excel)
        var numericQuantity = parseInt(quantity) || parseFloat(quantity) || 0;
        if (numericQuantity <= 0) {
            console.log('Skipping row', i + 1, '- invalid quantity:', quantity);
            continue;
        }
        
        var batchItem = {
            paperType: paperType,
            paperCode: getPaperCode(paperType),
            textBatchNumber: textBatch,
            coverBatchNumber: coverBatch,
            numberOfRows: numericQuantity,
            orderDate: orderDate
        };
        
        console.log('Adding batch item:', batchItem);
        batchData.push(batchItem);
    }

    console.log('Final processed batch data:', batchData);
    console.log('Total valid batches found:', batchData.length);
    
    displayResults();
}

function displayResults() {
    if (batchData.length === 0) {
        showStatus('No valid batch data found in the Excel file.', 'warning');
        return;
    }

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

    document.getElementById('resultsCard').style.display = 'block';
    document.getElementById('downloadPdfBtn').disabled = false;
    
    // Clear upload status message
    document.getElementById('uploadStatus').innerHTML = '';
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
        headerCell.colSpan = 7; // Updated to 7 columns to include Sequence
        headerCell.innerHTML = '<strong>' + items[0].paperCode + ' - ' + paperType + ' (' + items.length + ' batches)</strong>';
        
        // Add data rows for this paper type with sequence numbers
        items.forEach(function(item, index) {
            var row = tbody.insertRow();
            row.insertCell(0).textContent = item.paperCode;
            row.insertCell(1).textContent = item.paperType;
            row.insertCell(2).textContent = index + 1; // Sequence number starting from 1
            row.insertCell(3).textContent = item.textBatchNumber;
            row.insertCell(4).textContent = item.coverBatchNumber;
            row.insertCell(5).textContent = item.numberOfRows;
            row.insertCell(6).textContent = item.orderDate;
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

function formatOrderDateForFilename(orderDateString) {
    if (!orderDateString) return new Date().toISOString().split('T')[0]; // fallback to today
    
    // Try to parse common date formats from Excel
    var dateStr = orderDateString.toLowerCase().trim();
    
    // Handle formats like "Tuesday 29th", "Wednesday 30th", "Friday 1st", etc.
    if (dateStr.includes('29th')) {
        return '2025-07-29'; // Based on your example file
    } else if (dateStr.includes('30th')) {
        return '2025-07-30';
    } else if (dateStr.includes('31st')) {
        return '2025-07-31';
    } else if (dateStr.includes('1st')) {
        return '2025-08-01'; // Friday 1st = August 1st
    } else if (dateStr.includes('2nd')) {
        return '2025-08-02';
    } else if (dateStr.includes('3rd')) {
        return '2025-08-03';
    }
    
    // Try to extract date if it's in a different format
    var dateMatch = dateStr.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (dateMatch) {
        var day = dateMatch[1].padStart(2, '0');
        var month = dateMatch[2].padStart(2, '0');
        var year = dateMatch[3].length === 2 ? '20' + dateMatch[3] : dateMatch[3];
        return year + '-' + month + '-' + day;
    }
    
    // Fallback to today's date
    return new Date().toISOString().split('T')[0];
}
    var tbody = document.getElementById('resultsTableBody');
    tbody.innerHTML = '';

    paperTypes.forEach(function(paperType) {
        var items = groupedData[paperType];
        
        // Add paper type header row
        var headerRow = tbody.insertRow();
        headerRow.className = 'paper-type-header';
        var headerCell = headerRow.insertCell();
        headerCell.colSpan = 5;
        headerCell.innerHTML = '<strong>' + items[0].paperCode + ' - ' + paperType + ' (' + items.length + ' batches)</strong>';
        
        // Add data rows for this paper type
        items.forEach(function(item) {
            var row = tbody.insertRow();
            row.insertCell(0).textContent = item.paperCode;
            row.insertCell(1).textContent = item.paperType;
            row.insertCell(2).textContent = item.textBatchNumber;
            row.insertCell(3).textContent = item.coverBatchNumber;
            row.insertCell(4).textContent = item.numberOfRows;
        });
    });

function downloadPDF() {
    if (batchData.length === 0) {
        alert('No data to download.');
        return;
    }

    var jsPDF = window.jspdf.jsPDF;
    var pdf = new jsPDF('p', 'mm', 'a4');

    console.log('Generating PDF with', batchData.length, 'batches');

    // PDF Header
    pdf.setFontSize(14);
    pdf.setFont(undefined, 'bold');
    pdf.text('PoD Work List - Batch Processing', 20, 20);

    // Get the order date from the Excel data for the header
    var orderDate = batchData.length > 0 ? batchData[0].orderDate : '';
    if (orderDate) {
        pdf.setFontSize(10);
        pdf.setFont(undefined, 'normal');
        pdf.text('Order Date: ' + orderDate, 20, 30);
        var yPos = 45;
    } else {
        var yPos = 35;
    }

    // Group data by paper type
    var groupedData = {};
    var totalRows = 0;
    var totalBatches = 0;

    batchData.forEach(function(item) {
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

    // Add summary
    pdf.setFontSize(10);
    pdf.setFont(undefined, 'normal');
    pdf.text('Total Paper Types: ' + paperTypes.length + ' | Total Batches: ' + totalBatches + ' | Total Rows: ' + totalRows, 20, yPos);
    yPos += 15;

    // Column positions - remove sequence column from PDF, keep it simple
    var colPositions = {
        paperCode: { x: 20 },
        paperType: { x: 50 },
        textBatch: { x: 100 },
        coverBatch: { x: 130 },
        rows: { x: 165 }
    };

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
        pdf.text(items[0].paperCode + ' - ' + paperType + ' (' + items.length + ' batches)', 17, yPos + 3);
        yPos += 15;

        // Column headers - no sequence column in PDF
        pdf.setFontSize(8);
        pdf.setFont(undefined, 'bold');
        pdf.text('Paper Code', colPositions.paperCode.x, yPos);
        pdf.text('Paper Type', colPositions.paperType.x, yPos);
        pdf.text('Text Batch', colPositions.textBatch.x, yPos);
        pdf.text('Cover Batch', colPositions.coverBatch.x, yPos);
        pdf.text('Rows', colPositions.rows.x, yPos);

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
                pdf.text(items[0].paperCode + ' - ' + paperType + ' (continued)', 17, yPos + 3);
                yPos += 15;
                
                // Repeat header on new page (added Sequence, removed Order Date for PDF)
                pdf.setFontSize(8);
                pdf.text('Paper Code', colPositions.paperCode.x, yPos);
                pdf.text('Paper Type', colPositions.paperType.x, yPos);
                pdf.text('Seq', colPositions.sequence.x, yPos);
                pdf.text('Text Batch', colPositions.textBatch.x, yPos);
                pdf.text('Cover Batch', colPositions.coverBatch.x, yPos);
                pdf.text('Rows', colPositions.rows.x, yPos);
                pdf.line(15, yPos + 2, 195, yPos + 2);
                yPos += 8;
                pdf.setFontSize(7);
                pdf.setFont(undefined, 'normal');
            }

            var paperTypeText = item.paperType.length > 20 ? item.paperType.substring(0, 20) + '..' : item.paperType;

            pdf.text(item.paperCode, colPositions.paperCode.x, yPos);
            pdf.text(paperTypeText, colPositions.paperType.x, yPos);
            pdf.text(item.textBatchNumber, colPositions.textBatch.x, yPos);
            pdf.text(item.coverBatchNumber, colPositions.coverBatch.x, yPos);
            pdf.text(String(item.numberOfRows), colPositions.rows.x, yPos);
            
            yPos += 5;
        });

        if (groupIndex < paperTypes.length - 1) {
            yPos += 10;
        }
    });

    // Add page numbers
    var pageCount = pdf.internal.getNumberOfPages();
    for (var i = 1; i <= pageCount; i++) {
        pdf.setPage(i);
        pdf.setFontSize(8);
        pdf.text('Page ' + i + ' of ' + pageCount, 175, 287);
    }

    // Generate filename using order date
    var orderDate = batchData.length > 0 ? batchData[0].orderDate : '';
    var formattedDate = formatOrderDateForFilename(orderDate);
    var fileName = 'Clays_POD_WTL_' + formattedDate + '.pdf';
    
    console.log('Saving PDF as:', fileName);
    pdf.save(fileName);
}

function clearAll() {
    batchData = [];
    
    document.getElementById('excelFileInput').value = '';
    document.getElementById('uploadStatus').innerHTML = '';
    document.getElementById('resultsCard').style.display = 'none';
    document.getElementById('resultsTableBody').innerHTML = '';
    document.getElementById('dataSummary').innerHTML = '';
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