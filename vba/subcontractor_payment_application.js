// This code generates an Excel file with multiple sheets for a subcontractor payment application
import * as XLSX from 'xlsx';

// Create a new workbook
const workbook = XLSX.utils.book_new();

// ================ MAIN APPLICATION FORM ================
const applicationData = [
  ['SUBCONTRACTOR PAYMENT APPLICATION', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['Project Name:', '', '', 'Application No.:', '', '', ''],
  ['Project Address:', '', '', 'Period From:', '', '', ''],
  ['', '', '', 'Period To:', '', '', ''],
  ['Subcontractor:', '', '', 'Contract No.:', '', '', ''],
  ['Address:', '', '', 'Contract Date:', '', '', ''],
  ['Phone:', '', '', 'Contract Amount:', '$', '', ''],
  ['Email:', '', '', 'Change Orders:', '$', '', ''],
  ['', '', '', 'Revised Contract:', '$', '', ''],
  ['', '', '', '', '', '', ''],
  ['PAYMENT APPLICATION SUMMARY', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['', 'Original Contract', 'Change Orders', 'Revised Contract', 'Previous Billing', 'Current Billing', 'Total to Date'],
  ['1. General Conditions', '', '', '', '', '', ''],
  ['2. Site Work', '', '', '', '', '', ''],
  ['3. Concrete', '', '', '', '', '', ''],
  ['4. Masonry', '', '', '', '', '', ''],
  ['5. Metals', '', '', '', '', '', ''],
  ['6. Wood & Plastics', '', '', '', '', '', ''],
  ['7. Thermal & Moisture', '', '', '', '', '', ''],
  ['8. Doors & Windows', '', '', '', '', '', ''],
  ['9. Finishes', '', '', '', '', '', ''],
  ['10. Specialties', '', '', '', '', '', ''],
  ['11. Equipment', '', '', '', '', '', ''],
  ['12. Furnishings', '', '', '', '', '', ''],
  ['13. Special Construction', '', '', '', '', '', ''],
  ['14. Conveying Systems', '', '', '', '', '', ''],
  ['15. Mechanical', '', '', '', '', '', ''],
  ['16. Electrical', '', '', '', '', '', ''],
  ['17. Materials Stored', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['TOTALS', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['Less Retention (___%)', '', '', '', '', '', ''],
  ['Net Amount', '', '', '', '', '', ''],
  ['Less Previous Payments', '', '', '', '', '', ''],
  ['AMOUNT DUE THIS APPLICATION', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['CERTIFICATION', '', '', '', '', '', ''],
  ['The undersigned Subcontractor certifies that to the best of the Subcontractor\'s knowledge, information, and belief, the Work covered by this Application for Payment has been completed in accordance with the Contract Documents.', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['Subcontractor:', '', '', 'Date:', '', '', ''],
  ['By:', '', '', 'Title:', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['GENERAL CONTRACTOR APPROVAL', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['General Contractor:', '', '', 'Date:', '', '', ''],
  ['By:', '', '', 'Title:', '', '', '']
];

// Create the main application worksheet
const applicationWS = XLSX.utils.aoa_to_sheet(applicationData);

// Set column widths
applicationWS['!cols'] = [
  { wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 15 }
];

// Set some merged cells for headers
applicationWS['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }, // Title
  { s: { r: 11, c: 0 }, e: { r: 11, c: 6 } }, // Payment Application Summary
  { s: { r: 39, c: 0 }, e: { r: 39, c: 6 } }, // Certification text
  { s: { r: 3, c: 1 }, e: { r: 3, c: 2 } }, // Project Address
  { s: { r: 6, c: 1 }, e: { r: 6, c: 2 } }, // Subcontractor Address
];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, applicationWS, 'Payment Application');

// ================ LIEN WAIVER SHEET ================
const lienWaiverData = [
  ['CONDITIONAL WAIVER AND RELEASE OF LIENS', '', '', '', ''],
  ['', '', '', '', ''],
  ['Project Name:', '', '', '', ''],
  ['Project Address:', '', '', '', ''],
  ['Subcontractor:', '', '', '', ''],
  ['Invoice/Pay App #:', '', '', '', ''],
  ['Payment Amount:', '$', '', '', ''],
  ['', '', '', '', ''],
  ['CONDITIONAL WAIVER AND RELEASE', '', '', '', ''],
  ['', '', '', '', ''],
  ['Upon receipt of payment of the sum of $____________, the undersigned hereby waives and releases any and all rights to a mechanics lien, stop notice, or any right against a labor and material bond on the job, with respect to and on account of labor, services, equipment, or material furnished to the above referenced project through the date of ____________, except for disputed claims specifically described as follows:', '', '', '', ''],
  ['', '', '', '', ''],
  ['Disputed Claims (if any):', '', '', '', ''],
  ['', '', '', '', ''],
  ['', '', '', '', ''],
  ['', '', '', '', ''],
  ['This release covers a progress payment for work performed through the above date only and does not cover any retention, pending modifications and changes, or items furnished after that date.', '', '', '', ''],
  ['', '', '', '', ''],
  ['NOTICE: THIS DOCUMENT WAIVES RIGHTS UNCONDITIONALLY AND STATES THAT YOU HAVE BEEN PAID FOR GIVING UP THOSE RIGHTS. THIS DOCUMENT IS ENFORCEABLE AGAINST YOU IF YOU SIGN IT, EVEN IF YOU HAVE NOT BEEN PAID. IF YOU HAVE NOT BEEN PAID, USE A CONDITIONAL RELEASE FORM.', '', '', '', ''],
  ['', '', '', '', ''],
  ['Subcontractor:', '', '', '', ''],
  ['By:', '', '', 'Title:', ''],
  ['', '', '', '', ''],
  ['Date:', '', '', '', ''],
  ['', '', '', '', ''],
  ['State of:', '', '', '', ''],
  ['County of:', '', '', '', ''],
  ['', '', '', '', ''],
  ['On this ____ day of ____________, 20____, before me, the undersigned, a Notary Public in and for said State, personally appeared ________________________, known to me to be the person who executed the foregoing instrument.', '', '', '', ''],
  ['', '', '', '', ''],
  ['Notary Public:', '', '', '', ''],
  ['My Commission Expires:', '', '', '', '']
];

// Create the lien waiver worksheet
const lienWaiverWS = XLSX.utils.aoa_to_sheet(lienWaiverData);

// Set column widths
lienWaiverWS['!cols'] = [
  { wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }
];

// Set merged cells for headers and paragraphs
lienWaiverWS['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, // Title
  { s: { r: 8, c: 0 }, e: { r: 8, c: 4 } }, // Conditional Waiver and Release
  { s: { r: 10, c: 0 }, e: { r: 10, c: 4 } }, // Waiver text
  { s: { r: 16, c: 0 }, e: { r: 16, c: 4 } }, // Progress payment text
  { s: { r: 18, c: 0 }, e: { r: 18, c: 4 } }, // Notice text
  { s: { r: 28, c: 0 }, e: { r: 28, c: 4 } }, // Notary text
];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, lienWaiverWS, 'Lien Waiver');

// ================ TAX FORMS SHEET ================
const taxFormsData = [
  ['TAX FORMS CHECKLIST', '', '', ''],
  ['', '', '', ''],
  ['Subcontractor Name:', '', '', ''],
  ['Federal Tax ID (EIN):', '', '', ''],
  ['', '', '', ''],
  ['FORM VERIFICATION', 'FORM NUMBER', 'INCLUDED', 'EXPIRATION'],
  ['Federal W-9', 'IRS Form W-9', '', ''],
  ['State Tax Registration', '', '', ''],
  ['Local Business License', '', '', ''],
  ['Sales Tax Certificate', '', '', ''],
  ['Use Tax Declaration', '', '', ''],
  ['Workers Comp Exemption (if applicable)', '', '', ''],
  ['1099 Information (if applicable)', '', '', ''],
  ['State Contractor License', '', '', ''],
  ['', '', '', ''],
  ['NOTES:', '', '', ''],
  ['1. All tax forms must be current and valid for the contract period', '', '', ''],
  ['2. Any changes to tax status must be reported immediately', '', '', ''],
  ['3. Missing or expired tax forms may result in payment delays', '', '', ''],
  ['4. Federal W-9 must be updated annually', '', '', ''],
  ['', '', '', ''],
  ['CERTIFICATION', '', '', ''],
  ['', '', '', ''],
  ['I certify that all tax information provided is accurate and complete.', '', '', ''],
  ['', '', '', ''],
  ['Signature:', '', '', ''],
  ['Name (Print):', '', '', ''],
  ['Title:', '', '', ''],
  ['Date:', '', '', '']
];

// Create the tax forms worksheet
const taxFormsWS = XLSX.utils.aoa_to_sheet(taxFormsData);

// Set column widths
taxFormsWS['!cols'] = [
  { wch: 30 }, { wch: 20 }, { wch: 15 }, { wch: 15 }
];

// Set merged cells for headers
taxFormsWS['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }, // Title
  { s: { r: 21, c: 0 }, e: { r: 21, c: 3 } }, // Certification
  { s: { r: 23, c: 0 }, e: { r: 23, c: 3 } }, // Certification text
];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, taxFormsWS, 'Tax Forms');

// ================ MATERIALS STORED SHEET ================
const materialsStoredData = [
  ['MATERIALS STORED ON SITE', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['Project Name:', '', '', 'Application No.:', '', '', ''],
  ['Subcontractor:', '', '', 'Period Ending:', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['ITEM', 'DESCRIPTION', 'STORAGE LOCATION', 'DATE RECEIVED', 'QUANTITY', 'UNIT COST', 'TOTAL VALUE'],
  ['1', '', '', '', '', '', ''],
  ['2', '', '', '', '', '', ''],
  ['3', '', '', '', '', '', ''],
  ['4', '', '', '', '', '', ''],
  ['5', '', '', '', '', '', ''],
  ['6', '', '', '', '', '', ''],
  ['7', '', '', '', '', '', ''],
  ['8', '', '', '', '', '', ''],
  ['9', '', '', '', '', '', ''],
  ['10', '', '', '', '', '', ''],
  ['11', '', '', '', '', '', ''],
  ['12', '', '', '', '', '', ''],
  ['13', '', '', '', '', '', ''],
  ['14', '', '', '', '', '', ''],
  ['15', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['TOTAL MATERIALS STORED', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['CERTIFICATION', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['The undersigned Subcontractor certifies that the materials listed above have been purchased, delivered, and properly stored on site, protected from weather and damage, and are covered by insurance. The materials have not been included in previous payment applications, and proof of purchase is attached.', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['Subcontractor:', '', '', 'Date:', '', '', ''],
  ['By:', '', '', 'Title:', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['GENERAL CONTRACTOR VERIFICATION', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['The General Contractor confirms that the stored materials listed above have been verified for quantity and proper storage.', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['General Contractor:', '', '', 'Date:', '', '', ''],
  ['By:', '', '', 'Title:', '', '', '']
];

// Create the materials stored worksheet
const materialsStoredWS = XLSX.utils.aoa_to_sheet(materialsStoredData);

// Set column widths
materialsStoredWS['!cols'] = [
  { wch: 8 }, { wch: 25 }, { wch: 20 }, { wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 15 }
];

// Set merged cells for headers
materialsStoredWS['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }, // Title
  { s: { r: 24, c: 0 }, e: { r: 24, c: 6 } }, // Certification
  { s: { r: 26, c: 0 }, e: { r: 26, c: 6 } }, // Certification text
  { s: { r: 31, c: 0 }, e: { r: 31, c: 6 } }, // GC Verification
  { s: { r: 33, c: 0 }, e: { r: 33, c: 6 } }, // GC Verification text
];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, materialsStoredWS, 'Materials Stored');

// ================ CERTIFIED PAYROLL SHEET ================
const certifiedPayrollData = [
  ['CERTIFIED PAYROLL REPORT', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', ''],
  ['Project Name:', '', '', 'Payroll No.:', '', '', 'Week Ending:', '', '', '', '', ''],
  ['Project Address:', '', '', 'Contract No.:', '', '', '', '', '', '', '', ''],
  ['Subcontractor:', '', '', 'Federal ID No.:', '', '', '', '', '', '', '', ''],
  ['Address:', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', ''],
  ['EMPLOYEE INFORMATION', 'HOURS WORKED EACH DAY', '', '', '', '', '', 'COMPENSATION', '', '', 'DEDUCTIONS', ''],
  ['Name, Address, SSN (last 4)', 'Work Classification', 'S', 'M', 'T', 'W', 'T', 'F', 'S', 'Total Hours', 'Rate of Pay', 'Gross Amount', 'Federal Tax', 'State Tax', 'FICA', 'Other', 'Total Deductions', 'Net Pay'],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['CERTIFICATION', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', ''],
  ['I, _________________________ (Name and Title), do hereby state:', '', '', '', '', '', '', '', '', '', '', ''],
  ['(1) That I pay or supervise the payment of the persons employed by _________________________ (Subcontractor) on the _________________________ (Project); that during the payroll period commencing on the _____ day of _____________, 20___, and ending the _____ day of _____________, 20___, all persons employed on said project have been paid full weekly wages earned, that no rebates have been or will be made either directly or indirectly to or on behalf of said _________________________ (Subcontractor) from the full weekly wages earned by any person, and that no deductions have been made either directly or indirectly from the full wages earned by any person, other than permissible deductions as defined in applicable laws and regulations.', '', '', '', '', '', '', '', '', '', '', ''],
  ['(2) That any payrolls otherwise under this contract required to be submitted for the above period are correct and complete; that the wage rates for laborers or mechanics contained therein are not less than the applicable wage rates contained in any wage determination incorporated into the contract; and that the classifications set forth therein for each laborer or mechanic conform with the work performed.', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', ''],
  ['I declare under penalty of perjury that the foregoing is true and correct.', '', '', '', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', ''],
  ['Signature:', '', '', 'Date:', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '']
];

// Create the certified payroll worksheet
const certifiedPayrollWS = XLSX.utils.aoa_to_sheet(certifiedPayrollData);

// Set column widths (adjust as needed for the expanded columns)
certifiedPayrollWS['!cols'] = [
  { wch: 25 }, { wch: 15 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, 
  { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 12 }
];

// Set merged cells for headers and paragraphs
certifiedPayrollWS['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 17 } }, // Title
  { s: { r: 21, c: 0 }, e: { r: 21, c: 17 } }, // Certification
  { s: { r: 23, c: 0 }, e: { r: 23, c: 17 } }, // Name and Title line
  { s: { r: 24, c: 0 }, e: { r: 24, c: 17 } }, // Certification paragraph 1
  { s: { r: 25, c: 0 }, e: { r: 25, c: 17 } }, // Certification paragraph 2
  { s: { r: 27, c: 0 }, e: { r: 27, c: 17 } }, // Perjury statement
  { s: { r: 7, c: 2 }, e: { r: 7, c: 8 } }, // Hours worked heading
  { s: { r: 7, c: 9 }, e: { r: 7, c: 11 } }, // Compensation heading
  { s: { r: 7, c: 12 }, e: { r: 7, c: 16 } }, // Deductions heading
];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, certifiedPayrollWS, 'Certified Payroll');

// ================ INSURANCE CERTIFICATE SHEET ================
const insuranceData = [
  ['CERTIFICATE OF INSURANCE VERIFICATION', '', '', '', ''],
  ['', '', '', '', ''],
  ['Subcontractor:', '', '', '', ''],
  ['Project Name:', '', '', '', ''],
  ['', '', '', '', ''],
  ['INSURANCE POLICY INFORMATION', '', '', '', ''],
  ['', '', '', '', ''],
  ['INSURANCE TYPE', 'CARRIER', 'POLICY NUMBER', 'EXPIRATION DATE', 'COVERAGE LIMITS'],
  ['Commercial General Liability', '', '', '', ''],
  ['Workers Compensation', '', '', '', ''],
  ['Employer\'s Liability', '', '', '', ''],
  ['Automobile Liability', '', '', '', ''],
  ['Umbrella/Excess Liability', '', '', '', ''],
  ['Professional Liability (if applicable)', '', '', '', ''],
  ['Pollution Liability (if applicable)', '', '', '', ''],
  ['Builder\'s Risk (if applicable)', '', '', '', ''],
  ['', '', '', '', ''],
  ['REQUIRED INSURANCE VERIFICATION', '', '', '', ''],
  ['', '', '', '', ''],
  ['REQUIREMENT', 'MINIMUM REQUIRED', 'ACTUAL', 'COMPLIANT (Y/N)', 'NOTES'],
  ['General Liability - Each Occurrence', '$1,000,000', '', '', ''],
  ['General Liability - General Aggregate', '$2,000,000', '', '', ''],
  ['General Liability - Products/Completed Ops', '$2,000,000', '', '', ''],
  ['Auto Liability - Combined Single Limit', '$1,000,000', '', '', ''],
  ['Workers Compensation', 'Statutory Limits', '', '', ''],
  ['Employer\'s Liability', '$1,000,000', '', '', ''],
  ['Umbrella/Excess Liability', '$5,000,000', '', '', ''],
  ['', '', '', '', ''],
  ['ADDITIONAL INSURED VERIFICATION', '', '', '', ''],
  ['', '', '', '', ''],
  ['The following parties are listed as Additional Insured on the General Liability policy:', '', '', '', ''],
  ['1. General Contractor:', '', '', '', ''],
  ['2. Owner:', '', '', '', ''],
  ['3. Architect/Engineer:', '', '', '', ''],
  ['4. Others (as required):', '', '', '', ''],
  ['', '', '', '', ''],
  ['WAIVER OF SUBROGATION', '', '', '', ''],
  ['Waiver of Subrogation provided in favor of Additional Insureds?', 'YES / NO', '', '', ''],
  ['', '', '', '', ''],
  ['PRIMARY & NON-CONTRIBUTORY', '', '', '', ''],
  ['Coverage is Primary & Non-Contributory?', 'YES / NO', '', '', ''],
  ['', '', '', '', ''],
  ['NOTICE OF CANCELLATION', '', '', '', ''],
  ['30-day Notice of Cancellation provided?', 'YES / NO', '', '', ''],
  ['', '', '', '', ''],
  ['VERIFICATION', '', '', '', ''],
  ['', '', '', '', ''],
  ['I have reviewed the above insurance requirements and verify that all requirements have been met.', '', '', '', ''],
  ['', '', '', '', ''],
  ['Certificate Attached?', 'YES / NO', '', '', ''],
  ['', '', '', '', ''],
  ['Verified By:', '', '', '', ''],
  ['Title:', '', '', '', ''],
  ['Date:', '', '', '', '']
];

// Create the insurance certificate worksheet
const insuranceWS = XLSX.utils.aoa_to_sheet(insuranceData);

// Set column widths
insuranceWS['!cols'] = [
  { wch: 30 }, { wch: 20 }, { wch: 20 }, { wch: 15 }, { wch: 20 }
];

// Set merged cells for headers
insuranceWS['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, // Title
  { s: { r: 5, c: 0 }, e: { r: 5, c: 4 } }, // Insurance Policy Information
  { s: { r: 17, c: 0 }, e: { r: 17, c: 4 } }, // Required Insurance Verification
  { s: { r: 28, c: 0 }, e: { r: 28, c: 4 } }, // Additional Insured Verification
  { s: { r: 30, c: 0 }, e: { r: 30, c: 4 } }, // Additional Insured text
  { s: { r: 36, c: 0 }, e: { r: 36, c: 4 } }, // Waiver of Subrogation
  { s: { r: 39, c: 0 }, e: { r: 39, c: 4 } }, // Primary & Non-Contributory
  { s: { r: 42, c: 0 }, e: { r: 42, c: 4 } }, // Notice of Cancellation
  { s: { r: 45, c: 0 }, e: { r: 45, c: 4 } }, // Verification
  { s: { r: 47, c: 0 }, e: { r: 47, c: 4 } }, // Verification text
];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, insuranceWS, 'Insurance Certificate');

// Convert the workbook to a binary string
const excelBinaryString = XLSX.write(workbook, { type: 'binary', bookType: 'xlsx' });

// Helper function to convert string to ArrayBuffer
function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) {
    view[i] = s.charCodeAt(i) & 0xFF;
  }
  return buf;
}

// Create blob and download link
const blob = new Blob([s2ab(excelBinaryString)], { type: 'application/octet-stream' });
const url = URL.createObjectURL(blob);

// Create download link element
const downloadLink = document.createElement('a');
downloadLink.href = url;
downloadLink.download = 'Subcontractor_Payment_Application.xlsx';
downloadLink.textContent = 'Download Subcontractor Payment Application Excel File';
document.body.appendChild(downloadLink);

// Trigger click to start download
downloadLink.click();

// Clean up
URL.revokeObjectURL(url);
document.body.removeChild(downloadLink);

console.log('Excel file generated successfully!');
