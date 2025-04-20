// This code generates an Excel workbook with multiple sheets for construction submittal tracking
// It includes a submittal log, submittal package template, and submittal item form

import * as XLSX from 'xlsx';

// Create a new workbook
const workbook = XLSX.utils.book_new();

// SHEET 1: SUBMITTAL LOG
const submittalLogData = [
  ['CONSTRUCTION SUBMITTAL LOG', '', '', '', '', '', '', '', '', ''],
  ['Project Name:', 'ENTER PROJECT NAME', '', '', '', '', 'Project No.:', 'ENTER PROJECT NO.', '', ''],
  ['Contractor:', 'ENTER CONTRACTOR NAME', '', '', '', '', 'Updated:', new Date().toLocaleDateString(), '', ''],
  ['', '', '', '', '', '', '', '', '', ''],
  ['Submittal No.', 'Spec Section', 'Description', 'Submittal Type', 'Subcontractor', 'Required Date', 'Date Received', 'Status', 'Returned Date', 'Remarks'],
  ['001', '03 30 00', 'Concrete Mix Design', 'Product Data', 'ABC Concrete', '05/01/2025', '04/15/2025', 'Approved', '04/20/2025', ''],
  ['002', '04 20 00', 'Masonry Units', 'Samples', 'XYZ Masonry', '05/15/2025', '04/22/2025', 'Approved as Noted', '04/27/2025', 'Resubmit color samples'],
  ['003', '05 12 00', 'Structural Steel', 'Shop Drawings', 'Steel Fabricators Inc.', '05/30/2025', '', 'Pending', '', ''],
  ['004', '07 21 00', 'Insulation', 'Product Data', 'Insulation Co.', '06/15/2025', '', 'Pending', '', ''],
  ['005', '08 11 13', 'Hollow Metal Doors', 'Shop Drawings', 'Door Suppliers LLC', '06/30/2025', '', 'Pending', '', ''],
  ['', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '']
];

// Create worksheet and add data
const submittalLogWS = XLSX.utils.aoa_to_sheet(submittalLogData);

// Set column widths
submittalLogWS['!cols'] = [
  { width: 12 }, // Submittal No.
  { width: 15 }, // Spec Section
  { width: 25 }, // Description
  { width: 15 }, // Submittal Type
  { width: 20 }, // Subcontractor
  { width: 15 }, // Required Date
  { width: 15 }, // Date Received
  { width: 15 }, // Status
  { width: 15 }, // Returned Date
  { width: 25 }  // Remarks
];

// Add worksheet to workbook
XLSX.utils.book_append_sheet(workbook, submittalLogWS, 'Submittal Log');

// SHEET 2: SUBMITTAL PACKAGE TEMPLATE
const submittalPackageData = [
  ['SUBMITTAL PACKAGE COVER SHEET', '', '', '', '', ''],
  ['', '', '', '', '', ''],
  ['Project Name:', 'ENTER PROJECT NAME', '', 'Project No.:', 'ENTER PROJECT NO.', ''],
  ['Contractor:', 'ENTER CONTRACTOR NAME', '', 'Date:', new Date().toLocaleDateString(), ''],
  ['', '', '', '', '', ''],
  ['SUBMITTAL PACKAGE INFORMATION', '', '', '', '', ''],
  ['Package No.:', '', '', 'Specification Section(s):', '', ''],
  ['Package Description:', '', '', '', '', ''],
  ['Submitted By:', '', '', 'Title:', '', ''],
  ['', '', '', '', '', ''],
  ['SUBMITTAL ITEMS INCLUDED IN THIS PACKAGE', '', '', '', '', ''],
  ['Item No.', 'Description', 'Type', 'No. of Copies', 'Spec Reference', 'Remarks'],
  ['1', '', '', '', '', ''],
  ['2', '', '', '', '', ''],
  ['3', '', '', '', '', ''],
  ['4', '', '', '', '', ''],
  ['5', '', '', '', '', ''],
  ['', '', '', '', '', ''],
  ['REVIEW ACTION', '', '', '', '', ''],
  ['□ APPROVED', '□ APPROVED AS NOTED', '□ REVISE AND RESUBMIT', '□ REJECTED', '□ FOR INFORMATION ONLY', ''],
  ['', '', '', '', '', ''],
  ['Comments:', '', '', '', '', ''],
  ['', '', '', '', '', ''],
  ['', '', '', '', '', ''],
  ['', '', '', '', '', ''],
  ['Reviewed By:', '', '', 'Date:', '', ''],
  ['', '', '', '', '', '']
];

// Create worksheet and add data
const submittalPackageWS = XLSX.utils.aoa_to_sheet(submittalPackageData);

// Set column widths
submittalPackageWS['!cols'] = [
  { width: 15 },
  { width: 25 },
  { width: 20 },
  { width: 15 },
  { width: 20 },
  { width: 20 }
];

// Add worksheet to workbook
XLSX.utils.book_append_sheet(workbook, submittalPackageWS, 'Package Template');

// SHEET 3: SUBMITTAL ITEM FORM
const submittalItemData = [
  ['SUBMITTAL ITEM FORM', '', '', '', ''],
  ['', '', '', '', ''],
  ['Project Name:', 'ENTER PROJECT NAME', '', 'Project No.:', 'ENTER PROJECT NO.'],
  ['Contractor:', 'ENTER CONTRACTOR NAME', '', 'Date:', new Date().toLocaleDateString()],
  ['', '', '', '', ''],
  ['SUBMITTAL INFORMATION', '', '', '', ''],
  ['Submittal No.:', '', '', 'Spec Section:', ''],
  ['Description:', '', '', '', ''],
  ['Submittal Type:', '', '', '', ''],
  ['□ Product Data', '□ Shop Drawings', '□ Samples', '□ Quality Control', '□ Other:___________'],
  ['Subcontractor:', '', '', '', ''],
  ['Manufacturer:', '', '', '', ''],
  ['Supplier:', '', '', '', ''],
  ['', '', '', '', ''],
  ['CONTRACTOR REVIEW', '', '', '', ''],
  ['□ Approved', '□ Approved as Noted', '□ Revise and Resubmit', '□ Rejected', ''],
  ['Comments:', '', '', '', ''],
  ['', '', '', '', ''],
  ['', '', '', '', ''],
  ['Reviewer:', '', '', 'Date:', ''],
  ['', '', '', '', ''],
  ['ARCHITECT/ENGINEER REVIEW', '', '', '', ''],
  ['□ Approved', '□ Approved as Noted', '□ Revise and Resubmit', '□ Rejected', '□ For Information Only'],
  ['Comments:', '', '', '', ''],
  ['', '', '', '', ''],
  ['', '', '', '', ''],
  ['Reviewer:', '', '', 'Date:', ''],
  ['', '', '', '', '']
];

// Create worksheet and add data
const submittalItemWS = XLSX.utils.aoa_to_sheet(submittalItemData);

// Set column widths
submittalItemWS['!cols'] = [
  { width: 15 },
  { width: 20 },
  { width: 20 },
  { width: 15 },
  { width: 20 }
];

// Add worksheet to workbook
XLSX.utils.book_append_sheet(workbook, submittalItemWS, 'Item Form');

// SHEET 4: INSTRUCTIONS
const instructionsData = [
  ['CONSTRUCTION SUBMITTAL TRACKING WORKBOOK - INSTRUCTIONS', '', '', ''],
  ['', '', '', ''],
  ['Overview:', '', '', ''],
  ['This workbook contains templates for tracking construction submittals, creating submittal packages, and documenting individual submittal items.', '', '', ''],
  ['', '', '', ''],
  ['Sheets Included:', '', '', ''],
  ['1. Submittal Log - For tracking all project submittals', '', '', ''],
  ['2. Package Template - Template for creating submittal package cover sheets', '', '', ''],
  ['3. Item Form - Template for individual submittal item documentation', '', '', ''],
  ['', '', '', ''],
  ['How to Use:', '', '', ''],
  ['Submittal Log:', '', '', ''],
  ['- Enter project information at the top', '', '', ''],
  ['- Assign sequential submittal numbers', '', '', ''],
  ['- Enter specification section references', '', '', ''],
  ['- Track status and dates for each submittal', '', '', ''],
  ['- Update the log regularly as submittals are processed', '', '', ''],
  ['', '', '', ''],
  ['Package Template:', '', '', ''],
  ['- Complete this form when bundling multiple related submittals', '', '', ''],
  ['- Assign a package number', '', '', ''],
  ['- List all submittal items included in the package', '', '', ''],
  ['- Attach this as a cover sheet to the physical submittal package', '', '', ''],
  ['', '', '', ''],
  ['Item Form:', '', '', ''],
  ['- Complete this form for each individual submittal item', '', '', ''],
  ['- Document contractor review before submission', '', '', ''],
  ['- Record architect/engineer review after return', '', '', ''],
  ['- Include with the actual submittal materials', '', '', ''],
  ['', '', '', ''],
  ['Submittal Status Options:', '', '', ''],
  ['- Pending: Not yet received from subcontractor/supplier', '', '', ''],
  ['- In Review: Received but not yet reviewed', '', '', ''],
  ['- Approved: Reviewed and accepted without changes', '', '', ''],
  ['- Approved as Noted: Approved with minor changes noted', '', '', ''],
  ['- Revise and Resubmit: Requires correction and resubmission', '', '', ''],
  ['- Rejected: Not acceptable, must be resubmitted', '', '', ''],
  ['- For Information Only: Provided for reference, no approval required', '', '', ''],
  ['', '', '', ''],
  ['Tips:', '', '', ''],
  ['- Make a copy of the template sheets as needed', '', '', ''],
  ['- Maintain consistent submittal numbering', '', '', ''],
  ['- Use clear, descriptive titles for submittals', '', '', ''],
  ['- Track resubmittals with the original number plus a revision suffix (e.g., 001-R1)', '', '', ''],
  ['- Set up automatic conditional formatting to highlight late or overdue submittals', '', '', ''],
  ['', '', '', '']
];

// Create worksheet and add data
const instructionsWS = XLSX.utils.aoa_to_sheet(instructionsData);

// Set column widths
instructionsWS['!cols'] = [
  { width: 25 },
  { width: 30 },
  { width: 20 },
  { width: 20 }
];

// Add worksheet to workbook
XLSX.utils.book_append_sheet(workbook, instructionsWS, 'Instructions');

// Generate Excel file
const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
