import * as XLSX from 'xlsx';

/**
 * Export data to various formats
 */

const tryFormatDotNetDate = (value) => {
  if (!value) return value;

  // Handle .NET style JSON dates: /Date(1771899555000)/
  if (typeof value === 'string' && value.includes('/Date(')) {
    const match = value.match(/\/Date\((\d+)([+-]\d{4})?\)\//);
    if (match) {
      const timestamp = parseInt(match[1], 10);
      const date = new Date(timestamp);
      if (!Number.isNaN(date.getTime())) {
        const pad = (n) => String(n).padStart(2, '0');
        return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
      }
    }
  }

  // Handle Date objects (can appear when data comes from JS)
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const pad = (n) => String(n).padStart(2, '0');
    return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())} ${pad(value.getHours())}:${pad(value.getMinutes())}:${pad(value.getSeconds())}`;
  }

  return value;
};

// Excel sheet names must be <= 31 chars and cannot contain: \ / ? * [ ]
const sanitizeExcelSheetName = (name) => {
  const raw = String(name || 'Export');
  const cleaned = raw
    .replace(/[\\\/\?\*\[\]\:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const safe = cleaned || 'Export';
  return safe.length > 31 ? safe.slice(0, 31) : safe;
};

// Export to JSON
export const exportToJSON = (data, filename) => {
  const jsonData = JSON.stringify(data, null, 2);
  const blob = new Blob([jsonData], { type: 'application/json' });
  downloadBlob(blob, `${filename}.json`);
};

// Export to CSV
export const exportToCSV = (data, filename) => {
  if (!data || data.length === 0) {
    alert('No data to export');
    return;
  }

  // Get headers from first object
  const headers = Object.keys(data[0]);
  
  // Build CSV content
  const csvRows = [];
  
  // Add header row
  csvRows.push(headers.map(h => `"${h}"`).join(','));
  
  // Add data rows
  for (const row of data) {
    const values = headers.map(header => {
      let value = row[header];
      
      // Handle null/undefined
      if (value === null || value === undefined) {
        value = '';
      }
      
      // Handle arrays
      if (Array.isArray(value)) {
        value = value.map(tryFormatDotNetDate).join('; ');
      }
      
      // Handle objects
      if (typeof value === 'object' && value !== null) {
        const normalized = tryFormatDotNetDate(value);
        if (normalized !== value) value = normalized;
        value = JSON.stringify(value);
      }
      
      // Handle booleans
      if (typeof value === 'boolean') {
        value = value ? 'Yes' : 'No';
      }

      // Normalize .NET style dates if present
      value = tryFormatDotNetDate(value);
      
      // Convert to string and escape quotes
      value = String(value).replace(/"/g, '""');
      
      return `"${value}"`;
    });
    csvRows.push(values.join(','));
  }
  
  const csvContent = csvRows.join('\n');
  const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv; charset=utf-8' });
  downloadBlob(blob, `${filename}.csv`);
};

// Export to Excel
export const exportToExcel = (data, filename, sheetName = 'Export') => {
  if (!data || data.length === 0) {
    alert('No data to export');
    return;
  }

  // Process data - flatten arrays and handle booleans for Excel
  const processedData = data.map(row => {
    const processed = {};
    for (const [key, value] of Object.entries(row)) {
      if (Array.isArray(value)) {
        processed[key] = value.map(tryFormatDotNetDate).join('; ');
      } else if (typeof value === 'object' && value !== null) {
        const normalized = tryFormatDotNetDate(value);
        if (normalized !== value) {
          processed[key] = normalized;
        } else {
          processed[key] = JSON.stringify(value);
        }
      } else if (typeof value === 'boolean') {
        processed[key] = value ? 'Yes' : 'No';
      } else {
        processed[key] = tryFormatDotNetDate(value);
      }
    }
    return processed;
  });

  // Create workbook
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(processedData);
  
  // Auto-size columns (approximate)
  const headers = Object.keys(processedData[0]);
  const colWidths = headers.map((header) => {
    let maxWidth = header.length;
    processedData.forEach(row => {
      const cellValue = String(row[header] || '');
      maxWidth = Math.max(maxWidth, Math.min(cellValue.length, 50));
    });
    return { wch: maxWidth + 2 };
  });
  ws['!cols'] = colWidths;
  
  // Add worksheet to workbook
  XLSX.utils.book_append_sheet(wb, ws, sanitizeExcelSheetName(sheetName));
  
  // Save file
  XLSX.writeFile(wb, `${filename}.xlsx`);
};

// Export to PDF (using HTML table to print)
export const exportToPDF = (data, filename, title = 'Export Report') => {
  if (!data || data.length === 0) {
    alert('No data to export');
    return;
  }

  const headers = Object.keys(data[0]);
  
  // Create a new window for printing
  const printWindow = window.open('', '_blank');
  
  // Build HTML table
  let tableHTML = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>${title}</title>
      <style>
        body {
          font-family: Arial, sans-serif;
          padding: 20px;
          color: #333;
        }
        h1 {
          color: #1a5f7a;
          margin-bottom: 10px;
        }
        .meta {
          color: #666;
          font-size: 12px;
          margin-bottom: 20px;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          font-size: 11px;
        }
        th {
          background-color: #1a5f7a;
          color: white;
          padding: 10px 8px;
          text-align: left;
          border: 1px solid #ddd;
        }
        td {
          padding: 8px;
          border: 1px solid #ddd;
        }
        tr:nth-child(even) {
          background-color: #f9f9f9;
        }
        tr:hover {
          background-color: #f5f5f5;
        }
        @media print {
          body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
      </style>
    </head>
    <body>
      <h1>${title}</h1>
      <div class="meta">
        <p>Generated: ${new Date().toLocaleString()}</p>
        <p>Total Records: ${data.length}</p>
      </div>
      <table>
        <thead>
          <tr>
            ${headers.map(h => `<th>${formatHeader(h)}</th>`).join('')}
          </tr>
        </thead>
        <tbody>
  `;
  
  // Add data rows
  for (const row of data) {
    tableHTML += '<tr>';
    for (const header of headers) {
      let value = row[header];
      
      if (value === null || value === undefined) {
        value = '-';
      } else if (Array.isArray(value)) {
        value = value.map(tryFormatDotNetDate).join(', ');
      } else if (typeof value === 'object') {
        const normalized = tryFormatDotNetDate(value);
        if (normalized !== value) value = normalized;
        value = JSON.stringify(value);
      } else if (typeof value === 'boolean') {
        value = value ? 'Yes' : 'No';
      }

      value = tryFormatDotNetDate(value);
      
      tableHTML += `<td>${escapeHtml(String(value))}</td>`;
    }
    tableHTML += '</tr>';
  }
  
  tableHTML += `
        </tbody>
      </table>
      <script>
        window.onload = function() {
          window.print();
        }
      </script>
    </body>
    </html>
  `;
  
  printWindow.document.write(tableHTML);
  printWindow.document.close();
};

// Helper function to download blob
const downloadBlob = (blob, filename) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

// Helper function to format header names
const formatHeader = (header) => {
  return header
    .replace(/([A-Z])/g, ' $1')
    .replace(/^./, str => str.toUpperCase())
    .trim();
};

// Helper function to escape HTML
const escapeHtml = (text) => {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
};

// Main export function that handles different formats
export const handleExportWithFormat = (data, filename, format, title = 'Export Report') => {
  // Ensure data is always an array
  const safeData = Array.isArray(data) ? data : (data ? [data] : []);
  if (safeData.length === 0) {
    alert('No data to export');
    return;
  }

  const timestamp = new Date().toISOString().split('T')[0];
  const exportFilename = `${filename}_${timestamp}`;
  
  switch (format) {
    case 'json':
      exportToJSON(safeData, exportFilename);
      break;
    case 'csv':
      exportToCSV(safeData, exportFilename);
      break;
    case 'excel':
      exportToExcel(safeData, exportFilename, title);
      break;
    case 'pdf':
      exportToPDF(safeData, exportFilename, title);
      break;
    default:
      exportToJSON(safeData, exportFilename);
  }
};
