let data = []; // This holds the initial Excel data
let filteredData = []; // This holds the filtered data after user operations

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = '';

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null ? 'NULL' : cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Function to apply the operation
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.toUpperCase().trim();
    const operationColumns = document.getElementById('operation-columns').value.toUpperCase().trim().split(',');

    if (!primaryColumn || operationColumns.length === 0) {
        alert("Please fill in all fields.");
        return;
    }

    filteredData = data.filter(row => {
        const primaryValue = row[primaryColumn];
        return operationColumns.every(col => {
            const value = row[col];
            const isNullCheck = document.getElementById('operation').value === "null";
            const isAndOperation = document.getElementById('operation-type').value === "and";
            if (isNullCheck) {
                return isAndOperation ? value === null && primaryValue === null : value === null || primaryValue === null;
            } else {
                return isAndOperation ? value !== null && primaryValue !== null : value !== null || primaryValue !== null;
            }
        });
    });

    displaySheet(filteredData);
}

// Event listeners for button clicks
document.getElementById('apply-operation').addEventListener('click', applyOperation);

// Download button click event
document.getElementById('download-button').addEventListener('click', function() {
    document.getElementById('download-modal').style.display = 'flex';
});

// Close modal event
document.getElementById('close-modal').addEventListener('click', function() {
    document.getElementById('download-modal').style.display = 'none';
});

// Confirm download button
document.getElementById('confirm-download').addEventListener('click', function() {
    const filename = document.getElementById('filename').value || 'download';
    const format = document.getElementById('file-format').value;

    if (filteredData.length === 0) {
        alert("No data available to download.");
        return;
    }

    if (format === 'xlsx') {
        const ws = XLSX.utils.json_to_sheet(filteredData.map(row => {
            return Object.keys(row).reduce((acc, key) => {
                acc[key] = row[key] === null ? 'NULL' : row[key];
                return acc;
            }, {});
        }));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, `${filename}.xlsx`);
    } else if (format === 'csv') {
        const csvData = filteredData.map(row => Object.values(row).map(cell => cell === null ? 'NULL' : cell).join(',')).join('\n');
        const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.setAttribute('download', `${filename}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    } else if (format === 'jpg' || format === 'pdf') {
        alert("JPG and PDF download options require additional libraries.");
        // For JPG/PDF, consider using libraries like html2canvas and jsPDF
    }

    document.getElementById('download-modal').style.display = 'none';
});

// Load the Excel file on page load (replace 'your-file-url.xlsx' with the actual file URL)
loadExcelSheet('your-file-url.xlsx');
