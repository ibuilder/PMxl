// First, let's create a SheetJS workbook
import * as XLSX from 'xlsx';

// Create a new workbook
const wb = XLSX.utils.book_new();

// Define the header row for the permit log
const permitsHeader = [
  "Permit ID",
  "Project Name",
  "Project Address",
  "Permit Type",
  "Application Date",
  "Applicant Name",
  "Applicant Contact",
  "Reviewer Assigned",
  "Review Status",
  "Additional Info Requested",
  "Info Received Date",
  "Approval Date",
  "Permit Expiration",
  "Inspection Dates",
  "Inspection Results",
  "Final Approval Date",
  "Comments"
];

// Create sample data for demonstration (5 rows of example data)
const samplePermitData = [
  [
    "P-2025-001", 
    "Main Street Renovation", 
    "123 Main St, Anytown", 
    "Building Alteration", 
    "01/15/2025", 
    "John Smith", 
    "john.smith@email.com", 
    "Mary Johnson", 
    "Approved", 
    "No", 
    "", 
    "02/05/2025", 
    "08/05/2025", 
    "02/28/2025, 03/15/2025", 
    "Pass, Pass", 
    "03/20/2025", 
    "Project completed ahead of schedule"
  ],
  [
    "P-2025-002", 
    "Oak Plaza Development", 
    "456 Oak Ave, Anytown", 
    "New Construction", 
    "02/03/2025", 
    "Sarah Williams", 
    "swilliams@construction.com", 
    "Tom Lee", 
    "In Review", 
    "Yes", 
    "03/01/2025", 
    "", 
    "", 
    "", 
    "", 
    "", 
    "Structural calculations requested"
  ],
  [
    "P-2025-003", 
    "Elm Street Roofing", 
    "789 Elm St, Anytown", 
    "Roof Repair", 
    "02/10/2025", 
    "Robert Johnson", 
    "rjohnson@roofers.com", 
    "Linda Chen", 
    "Approved", 
    "No", 
    "", 
    "02/25/2025", 
    "08/25/2025", 
    "03/10/2025", 
    "Pass", 
    "03/12/2025", 
    ""
  ],
  [
    "P-2025-004", 
    "Maple Court Addition", 
    "321 Maple Ct, Anytown", 
    "Building Addition", 
    "02/18/2025", 
    "David Miller", 
    "dmiller@homes.net", 
    "Mary Johnson", 
    "Additional Info Required", 
    "Yes", 
    "", 
    "", 
    "", 
    "", 
    "", 
    "", 
    "Zoning variance needed"
  ],
  [
    "P-2025-005", 
    "Pine Street Demolition", 
    "567 Pine St, Anytown", 
    "Demolition", 
    "03/01/2025", 
    "Michael Brown", 
    "mbrown@demo.com", 
    "Tom Lee", 
    "Approved", 
    "No", 
    "", 
    "03/15/2025", 
    "09/15/2025", 
    "", 
    "", 
    "", 
    "Environmental clearance verified"
  ]
];

// Convert the header and data to a worksheet
const ws_data = [permitsHeader, ...samplePermitData];
const ws = XLSX.utils.aoa_to_sheet(ws_data);

// Set column widths for better readability
ws['!cols'] = [
  { width: 10 },  // Permit ID
  { width: 25 },  // Project Name
  { width: 30 },  // Project Address
  { width: 20 },  // Permit Type
  { width: 15 },  // Application Date
  { width: 20 },  // Applicant Name
  { width: 25 },  // Applicant Contact
  { width: 20 },  // Reviewer Assigned
  { width: 20 },  // Review Status
  { width: 15 },  // Additional Info Requested
  { width: 15 },  // Info Received Date
  { width: 15 },  // Approval Date
  { width: 15 },  // Permit Expiration
  { width: 20 },  // Inspection Dates
  { width: 20 },  // Inspection Results
  { width: 15 },  // Final Approval Date
  { width: 40 },  // Comments
];

// Style the header row
for (let i = 0; i < permitsHeader.length; i++) {
  const cellRef = XLSX.utils.encode_cell({ r: 0, c: i });
  if (!ws[cellRef]) ws[cellRef] = {};
  ws[cellRef].s = {
    fill: { fgColor: { rgb: "DDDDDD" } },
    font: { bold: true }
  };
}

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(wb, ws, "Permit Log");

// Create a second sheet for permit statistics
const statsHeader = [
  "Metric",
  "Count",
  "Percentage"
];

const statsData = [
  ["Total Permits", 5, "100%"],
  ["Approved Permits", 3, "60%"],
  ["Pending Review", 1, "20%"],
  ["Additional Info Required", 1, "20%"],
  ["Expired Permits", 0, "0%"]
];

const stats_ws_data = [statsHeader, ...statsData];
const stats_ws = XLSX.utils.aoa_to_sheet(stats_ws_data);

// Set column widths for stats sheet
stats_ws['!cols'] = [
  { width: 25 },  // Metric
  { width: 10 },  // Count
  { width: 15 },  // Percentage
];

// Style the header row
for (let i = 0; i < statsHeader.length; i++) {
  const cellRef = XLSX.utils.encode_cell({ r: 0, c: i });
  if (!stats_ws[cellRef]) stats_ws[cellRef] = {};
  stats_ws[cellRef].s = {
    fill: { fgColor: { rgb: "DDDDDD" } },
    font: { bold: true }
  };
}

// Add the stats worksheet to the workbook
XLSX.utils.book_append_sheet(wb, stats_ws, "Statistics");

// Generate the Excel file
const excelData = XLSX.write(wb, { type: "array", bookType: "xlsx" });

// For demonstration purposes, you would save this to a file in a real environment
// But in this case, we're just showing the structure that would be created
console.log("Excel file created with sheets: " + wb.SheetNames.join(", "));

// Return an object describing what we've created
return {
  filename: "Construction_Permitting_Log.xlsx",
  sheets: wb.SheetNames,
  permitLogColumns: permitsHeader.length,
  permitLogRows: samplePermitData.length + 1,
  structure: {
    "Permit Log": {
      columns: permitsHeader,
      rowCount: samplePermitData.length + 1
    },
    "Statistics": {
      columns: statsHeader,
      rowCount: statsData.length + 1
    }
  }
};
