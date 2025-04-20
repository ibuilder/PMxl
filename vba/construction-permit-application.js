// Create a permit application form template in Excel
import * as XLSX from 'xlsx';

// Create a new workbook
const wb = XLSX.utils.book_new();

// Define the sections and fields for the application form
const applicationForm = [
  // Title and Instructions
  ["CONSTRUCTION PERMIT APPLICATION FORM"],
  ["Please complete all sections of this form. Fields marked with an asterisk (*) are required."],
  [""],
  
  // Section 1: Applicant Information
  ["SECTION 1: APPLICANT INFORMATION"],
  ["*Applicant Name:", ""],
  ["Company Name:", ""],
  ["*Mailing Address:", ""],
  ["*City:", "", "*State:", "", "*ZIP Code:", ""],
  ["*Phone Number:", "", "Email:", ""],
  ["Contractor License #:", "", "License Type:", ""],
  [""],
  
  // Section 2: Project Information
  ["SECTION 2: PROJECT INFORMATION"],
  ["*Project Address:", ""],
  ["*City:", "", "*State:", "", "*ZIP Code:", ""],
  ["*APN/Parcel Number:", ""],
  ["Zoning District:", ""],
  ["*Estimated Project Cost: $", ""],
  ["*Project Description:", ""],
  ["", ""],  // Extra space for description
  ["", ""],  // Extra space for description
  [""],
  
  // Section 3: Project Type
  ["SECTION 3: PROJECT TYPE (Check all that apply)"],
  ["[ ] New Construction", "", "[ ] Renovation/Remodel", ""],
  ["[ ] Addition", "", "[ ] Demolition", ""],
  ["[ ] Plumbing", "", "[ ] Electrical", ""],
  ["[ ] Mechanical", "", "[ ] Roofing", ""],
  ["[ ] Grading/Excavation", "", "[ ] Other (specify):", ""],
  [""],
  
  // Section 4: Building Information
  ["SECTION 4: BUILDING INFORMATION"],
  ["*Occupancy Type:", ""],
  ["*Construction Type:", ""],
  ["*Number of Stories:", "", "*Building Height:", ""],
  ["*Total Square Footage:", ""],
  ["Existing Square Footage:", "", "New Square Footage:", ""],
  ["Number of Bedrooms:", "", "Number of Bathrooms:", ""],
  ["Fire Sprinklers: [ ] Yes  [ ] No", ""],
  [""],
  
  // Section 5: Required Attachments
  ["SECTION 5: REQUIRED ATTACHMENTS"],
  ["Check all items included with this application:"],
  ["[ ] Site Plan", "", "[ ] Floor Plan", ""],
  ["[ ] Elevation Drawings", "", "[ ] Structural Calculations", ""],
  ["[ ] Environmental Documents", "", "[ ] Title 24 Energy Calculations", ""],
  ["[ ] Soils Report", "", "[ ] Other:", ""],
  [""],
  
  // Section 6: Declarations and Signatures
  ["SECTION 6: DECLARATIONS AND SIGNATURES"],
  ["I hereby certify that I have read and examined this application and know the same to be true and correct. All provisions of laws and ordinances governing this type of work will be complied with whether specified herein or not. The granting of a permit does not presume to give authority to violate or cancel the provisions of any federal, state, or local law regulating construction or the performance of construction."],
  [""],
  ["*Owner/Authorized Agent Signature:", "", "*Date:", ""],
  ["*Print Name:", "", "*Title:", ""],
  [""],
  
  // For Official Use Only
  ["FOR OFFICIAL USE ONLY"],
  ["Permit #:", "", "Date Received:", ""],
  ["Received By:", "", "Fee Amount: $", ""],
  ["Review Required: [ ] Building  [ ] Planning  [ ] Engineering  [ ] Fire  [ ] Health"],
  ["Comments:", ""],
  ["", ""],  // Extra space for comments
  ["", ""]   // Extra space for comments
];

// Create the worksheet
const ws = XLSX.utils.aoa_to_sheet(applicationForm);

// Set column widths
ws['!cols'] = [
  { width: 30 },
  { width: 25 },
  { width: 20 },
  { width: 25 }
];

// Set row heights for title and section headers
ws['!rows'] = [];
for (let i = 0; i < applicationForm.length; i++) {
  if (i === 0) {
    // Title row
    ws['!rows'][i] = { hpt: 30 }; // Height in points
  } else if (applicationForm[i][0].startsWith("SECTION")) {
    // Section headers
    ws['!rows'][i] = { hpt: 22 };
  } else {
    // Regular rows
    ws['!rows'][i] = { hpt: 18 };
  }
}

// Style the cells
// First, let's style the title
const titleCell = XLSX.utils.encode_cell({ r: 0, c: 0 });
ws[titleCell].s = {
  font: { bold: true, sz: 16 },
  alignment: { horizontal: 'center' }
};
// Merge the title across columns
ws['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }
];

// Style section headers
for (let i = 0; i < applicationForm.length; i++) {
  if (applicationForm[i][0].startsWith("SECTION") || applicationForm[i][0] === "FOR OFFICIAL USE ONLY") {
    const cellRef = XLSX.utils.encode_cell({ r: i, c: 0 });
    ws[cellRef].s = {
      font: { bold: true, sz: 12 },
      fill: { fgColor: { rgb: "DDDDDD" } }
    };
    
    // Merge section headers across columns
    ws['!merges'].push({ s: { r: i, c: 0 }, e: { r: i, c: 3 } });
  }
}

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(wb, ws, "Permit Application");

// Generate the Excel file
const excelData = XLSX.write(wb, { type: "array", bookType: "xlsx" });

// For demonstration purposes, you would save this to a file in a real environment
console.log("Permit Application Form created");

// Return an object describing what we've created
return {
  filename: "Construction_Permit_Application.xlsx",
  sheets: wb.SheetNames,
  rowCount: applicationForm.length,
  sections: [
    "Applicant Information",
    "Project Information",
    "Project Type",
    "Building Information",
    "Required Attachments",
    "Declarations and Signatures",
    "Official Use Only"
  ]
};
