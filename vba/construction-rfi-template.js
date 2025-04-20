// This code will generate an Excel file with two sheets:
// 1. RFI Log - To track all RFIs
// 2. RFI Form - Template form for new RFIs

// This is meant to be run in a Node.js environment with the exceljs library
// You would need to: npm install exceljs

const ExcelJS = require('exceljs');
const workbook = new ExcelJS.Workbook();

// Add metadata
workbook.creator = 'RFI System';
workbook.lastModifiedBy = 'Project Manager';
workbook.created = new Date();
workbook.modified = new Date();

// ===== SHEET 1: RFI LOG =====
const logSheet = workbook.addWorksheet('RFI Log', {
  properties: { tabColor: { argb: '4472C4' } }
});

// Define columns
logSheet.columns = [
  { header: 'RFI #', key: 'id', width: 8 },
  { header: 'Date Submitted', key: 'dateSubmitted', width: 15 },
  { header: 'Title', key: 'title', width: 30 },
  { header: 'Description', key: 'description', width: 50 },
  { header: 'Requested By', key: 'requestedBy', width: 20 },
  { header: 'Company', key: 'company', width: 20 },
  { header: 'Discipline', key: 'discipline', width: 15 },
  { header: 'Priority', key: 'priority', width: 10 },
  { header: 'Status', key: 'status', width: 12 },
  { header: 'Assigned To', key: 'assignedTo', width: 20 },
  { header: 'Response Date', key: 'responseDate', width: 15 },
  { header: 'Days Open', key: 'daysOpen', width: 10 },
  { header: 'File Attachments', key: 'attachments', width: 30 }
];

// Style the header row
logSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
logSheet.getRow(1).fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '4472C4' }
};
logSheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

// Add some sample data
logSheet.addRow({
  id: 'RFI-001',
  dateSubmitted: new Date(),
  title: 'Foundation Depth Clarification',
  description: 'Need confirmation on the foundation depth for column locations B3 through B8',
  requestedBy: 'John Smith',
  company: 'ABC Construction',
  discipline: 'Structural',
  priority: 'High',
  status: 'Open',
  assignedTo: 'Jane Architect',
  responseDate: '',
  daysOpen: '=IF(K2="",TODAY()-B2,K2-B2)',
  attachments: 'Foundation-plan-rev2.pdf'
});

// Create a table for easy filtering
logSheet.addTable({
  name: 'RFILog',
  ref: 'A1',
  headerRow: true,
  style: {
    theme: 'TableStyleMedium2',
    showRowStripes: true,
  },
  columns: [
    { name: 'RFI #' },
    { name: 'Date Submitted' },
    { name: 'Title' },
    { name: 'Description' },
    { name: 'Requested By' },
    { name: 'Company' },
    { name: 'Discipline' },
    { name: 'Priority' },
    { name: 'Status' },
    { name: 'Assigned To' },
    { name: 'Response Date' },
    { name: 'Days Open' },
    { name: 'File Attachments' }
  ],
  rows: []
});

// Create data validation for Status column
logSheet.dataValidations.add('I2:I1000', {
  type: 'list',
  allowBlank: false,
  formulae: ['"Open,In Review,Responded,Closed"']
});

// Create data validation for Priority column
logSheet.dataValidations.add('H2:H1000', {
  type: 'list',
  allowBlank: false,
  formulae: ['"Low,Medium,High,Critical"']
});

// Create data validation for Discipline column
logSheet.dataValidations.add('G2:G1000', {
  type: 'list',
  allowBlank: false,
  formulae: ['"Architectural,Structural,Mechanical,Electrical,Plumbing,Civil,Other"']
});

// ===== SHEET 2: RFI FORM =====
const formSheet = workbook.addWorksheet('RFI Form', {
  properties: { tabColor: { argb: '70AD47' } }
});

// Set some column widths
formSheet.columns = [
  { width: 20 }, // A
  { width: 50 }, // B
  { width: 20 }, // C
  { width: 20 }  // D
];

// Add company logo placeholder
formSheet.mergeCells('A1:B3');
formSheet.getCell('A1').value = 'COMPANY LOGO';
formSheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
formSheet.getCell('A1').font = { bold: true, size: 16 };
formSheet.getCell('A1').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'E8E8E8' }
};

// RFI Form Title
formSheet.mergeCells('C1:D3');
formSheet.getCell('C1').value = 'REQUEST FOR INFORMATION';
formSheet.getCell('C1').alignment = { vertical: 'middle', horizontal: 'center' };
formSheet.getCell('C1').font = { bold: true, size: 16 };
formSheet.getCell('C1').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '70AD47' }
};

// Add horizontal separator
formSheet.mergeCells('A4:D4');
formSheet.getRow(4).height = 5;
formSheet.getCell('A4').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '000000' }
};

// RFI Details Section
formSheet.mergeCells('A5:D5');
formSheet.getCell('A5').value = 'RFI DETAILS';
formSheet.getCell('A5').font = { bold: true, size: 12 };
formSheet.getCell('A5').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'D0D0D0' }
};

// RFI Metadata
formSheet.getCell('A6').value = 'RFI Number:';
formSheet.getCell('A6').font = { bold: true };
formSheet.getCell('B6').value = 'RFI-XXX';

formSheet.getCell('C6').value = 'Date Submitted:';
formSheet.getCell('C6').font = { bold: true };
formSheet.getCell('D6').value = new Date();
formSheet.getCell('D6').numFmt = 'mm/dd/yyyy';

formSheet.getCell('A7').value = 'Project:';
formSheet.getCell('A7').font = { bold: true };
formSheet.getCell('B7').value = '';

formSheet.getCell('C7').value = 'Priority:';
formSheet.getCell('C7').font = { bold: true };

// Priority dropdown
formSheet.dataValidations.add('D7', {
  type: 'list',
  allowBlank: false,
  formulae: ['"Low,Medium,High,Critical"']
});

formSheet.getCell('A8').value = 'Requested By:';
formSheet.getCell('A8').font = { bold: true };
formSheet.getCell('B8').value = '';

formSheet.getCell('C8').value = 'Company:';
formSheet.getCell('C8').font = { bold: true };
formSheet.getCell('D8').value = '';

formSheet.getCell('A9').value = 'Discipline:';
formSheet.getCell('A9').font = { bold: true };

// Discipline dropdown
formSheet.dataValidations.add('B9', {
  type: 'list',
  allowBlank: false,
  formulae: ['"Architectural,Structural,Mechanical,Electrical,Plumbing,Civil,Other"']
});

formSheet.getCell('C9').value = 'Required Date:';
formSheet.getCell('C9').font = { bold: true };
formSheet.getCell('D9').value = '';
formSheet.getCell('D9').numFmt = 'mm/dd/yyyy';

// Request Section
formSheet.mergeCells('A10:D10');
formSheet.getCell('A10').value = 'REQUEST';
formSheet.getCell('A10').font = { bold: true, size: 12 };
formSheet.getCell('A10').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'D0D0D0' }
};

formSheet.getCell('A11').value = 'Request Title:';
formSheet.getCell('A11').font = { bold: true };
formSheet.mergeCells('B11:D11');

formSheet.getCell('A12').value = 'Specification Reference:';
formSheet.getCell('A12').font = { bold: true };
formSheet.mergeCells('B12:D12');

formSheet.getCell('A13').value = 'Drawing Reference:';
formSheet.getCell('A13').font = { bold: true };
formSheet.mergeCells('B13:D13');

formSheet.getCell('A14').value = 'Description of Request:';
formSheet.getCell('A14').font = { bold: true };
formSheet.mergeCells('A15:D19');
formSheet.getCell('A15').alignment = { wrapText: true, vertical: 'top' };
formSheet.getCell('A15').border = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' }
};

formSheet.getCell('A20').value = 'Attachments:';
formSheet.getCell('A20').font = { bold: true };
formSheet.mergeCells('B20:D20');

// Response Section
formSheet.mergeCells('A21:D21');
formSheet.getCell('A21').value = 'RESPONSE (To be completed by Design Team)';
formSheet.getCell('A21').font = { bold: true, size: 12 };
formSheet.getCell('A21').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'D0D0D0' }
};

formSheet.getCell('A22').value = 'Responded By:';
formSheet.getCell('A22').font = { bold: true };
formSheet.getCell('B22').value = '';

formSheet.getCell('C22').value = 'Response Date:';
formSheet.getCell('C22').font = { bold: true };
formSheet.getCell('D22').value = '';
formSheet.getCell('D22').numFmt = 'mm/dd/yyyy';

formSheet.mergeCells('A23:D23');
formSheet.getCell('A23').value = 'Response:';
formSheet.getCell('A23').font = { bold: true };

formSheet.mergeCells('A24:D28');
formSheet.getCell('A24').alignment = { wrapText: true, vertical: 'top' };
formSheet.getCell('A24').border = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' }
};

formSheet.getCell('A29').value = 'Response Attachments:';
formSheet.getCell('A29').font = { bold: true };
formSheet.mergeCells('B29:D29');

// Additional Comments Section
formSheet.mergeCells('A30:D30');
formSheet.getCell('A30').value = 'ADDITIONAL COMMENTS';
formSheet.getCell('A30').font = { bold: true, size: 12 };
formSheet.getCell('A30').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'D0D0D0' }
};

formSheet.mergeCells('A31:D34');
formSheet.getCell('A31').alignment = { wrapText: true, vertical: 'top' };
formSheet.getCell('A31').border = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' }
};

// Add instructions
formSheet.mergeCells('A36:D37');
formSheet.getCell('A36').value = 'INSTRUCTIONS: Complete all fields in the RFI DETAILS and REQUEST sections. Submit to the architect/engineer for response. Use the log sheet to track all RFIs.';
formSheet.getCell('A36').alignment = { wrapText: true };
formSheet.getCell('A36').font = { italic: true };

// Save the file
// In a real environment, you would do:
// await workbook.xlsx.writeFile('Construction_RFI_Template.xlsx');

// For demonstration, this code shows how the Excel file would be structured
// To use this template, copy this code into a Node.js environment with ExcelJS installed
