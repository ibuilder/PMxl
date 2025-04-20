// This code creates an Excel workbook with multiple sheets for a construction site safety program
// Using SheetJS (xlsx) library for Excel manipulation

// Create a new workbook
const workbook = XLSX.utils.book_new();

// 1. EMPLOYEE ORIENTATION FORM
const orientationData = [
  ["CONSTRUCTION SITE SAFETY PROGRAM", "", "", "", ""],
  ["EMPLOYEE ORIENTATION FORM", "", "", "", ""],
  ["", "", "", "", ""],
  ["Project Name:", "", "Project #:", "", ""],
  ["Project Location:", "", "", "", ""],
  ["Employee Name:", "", "Employee ID:", "", ""],
  ["Job Title:", "", "Start Date:", "", ""],
  ["Supervisor:", "", "", "", ""],
  ["", "", "", "", ""],
  ["ORIENTATION CHECKLIST", "YES", "NO", "N/A", "COMMENTS"],
  ["Company Safety Policy Reviewed", "", "", "", ""],
  ["Safety Rules and Procedures Explained", "", "", "", ""],
  ["Personal Protective Equipment Requirements", "", "", "", ""],
  ["Emergency Procedures and Exit Routes", "", "", "", ""],
  ["First Aid Locations and Procedures", "", "", "", ""],
  ["Incident Reporting Procedures", "", "", "", ""],
  ["Hazard Communication Program", "", "", "", ""],
  ["Confined Space Procedures", "", "", "", ""],
  ["Fall Protection Requirements", "", "", "", ""],
  ["Lockout/Tagout Procedures", "", "", "", ""],
  ["Fire Prevention and Protection", "", "", "", ""],
  ["Electrical Safety", "", "", "", ""],
  ["Hand and Power Tool Safety", "", "", "", ""],
  ["Heavy Equipment Safety", "", "", "", ""],
  ["Scaffolding and Ladder Safety", "", "", "", ""],
  ["", "", "", "", ""],
  ["Employee Comments/Questions:", "", "", "", ""],
  ["", "", "", "", ""],
  ["", "", "", "", ""],
  ["I acknowledge that I have received orientation on the above items and understand the safety requirements for this project.", "", "", "", ""],
  ["", "", "", "", ""],
  ["Employee Signature:", "", "Date:", "", ""],
  ["Orientation Conducted By:", "", "Date:", "", ""],
  ["Supervisor Signature:", "", "Date:", "", ""]
];

// Create the orientation worksheet
const orientationWS = XLSX.utils.aoa_to_sheet(orientationData);

// Add column widths and formatting
orientationWS['!cols'] = [
  { wch: 40 }, // A
  { wch: 10 }, // B
  { wch: 10 }, // C
  { wch: 10 }, // D
  { wch: 30 }  // E
];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, orientationWS, "Employee Orientation");

// 2. JOB HAZARD ANALYSIS FORM
const jhaData = [
  ["JOB HAZARD ANALYSIS (JHA) FORM", "", "", "", ""],
  ["", "", "", "", ""],
  ["Project Name:", "", "Project #:", "", ""],
  ["Location:", "", "Date:", "", ""],
  ["Task Description:", "", "", "", ""],
  ["Prepared By:", "", "Reviewed By:", "", ""],
  ["Required PPE:", "", "", "", ""],
  ["Required Tools/Equipment:", "", "", "", ""],
  ["Required Training:", "", "", "", ""],
  ["", "", "", "", ""],
  ["TASK STEPS", "POTENTIAL HAZARDS", "RISK LEVEL (H/M/L)", "CONTROL MEASURES", "RESPONSIBLE PERSON"],
  ["1.", "", "", "", ""],
  ["", "", "", "", ""],
  ["2.", "", "", "", ""],
  ["", "", "", "", ""],
  ["3.", "", "", "", ""],
  ["", "", "", "", ""],
  ["4.", "", "", "", ""],
  ["", "", "", "", ""],
  ["5.", "", "", "", ""],
  ["", "", "", "", ""],
  ["6.", "", "", "", ""],
  ["", "", "", "", ""],
  ["7.", "", "", "", ""],
  ["", "", "", "", ""],
  ["8.", "", "", "", ""],
  ["", "", "", "", ""],
  ["", "", "", "", ""],
  ["Emergency Procedures:", "", "", "", ""],
  ["", "", "", "", ""],
  ["APPROVALS", "", "", "", ""],
  ["Supervisor:", "", "Date:", "", ""],
  ["Safety Representative:", "", "Date:", "", ""],
  ["Project Manager:", "", "Date:", "", ""]
];

// Create the JHA worksheet
const jhaWS = XLSX.utils.aoa_to_sheet(jhaData);

// Add column widths
jhaWS['!cols'] = [
  { wch: 30 }, // A
  { wch: 30 }, // B
  { wch: 15 }, // C
  { wch: 30 }, // D
  { wch: 20 }  // E
];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, jhaWS, "Job Hazard Analysis");

// 3. PRETASK PLAN FORM
const pretaskData = [
  ["DAILY PRE-TASK PLAN", "", "", "", ""],
  ["", "", "", "", ""],
  ["Project Name:", "", "Project #:", "", ""],
  ["Task Location:", "", "Date:", "", ""],
  ["Task Description:", "", "", "", ""],
  ["Supervisor:", "", "", "", ""],
  ["", "", "", "", ""],
  ["CREW MEMBERS (Print Name & Sign)", "", "", "", ""],
  ["1.", "", "5.", "", ""],
  ["2.", "", "6.", "", ""],
  ["3.", "", "7.", "", ""],
  ["4.", "", "8.", "", ""],
  ["", "", "", "", ""],
  ["HAZARD IDENTIFICATION & CONTROL MEASURES", "", "", "", ""],
  ["POTENTIAL HAZARDS", "YES/NO", "CONTROL MEASURES", "", ""],
  ["Falls > 6 feet", "", "", "", ""],
  ["Electrical Hazards", "", "", "", ""],
  ["Confined Spaces", "", "", "", ""],
  ["Hot Work", "", "", "", ""],
  ["Hazardous Materials", "", "", "", ""],
  ["Heavy Equipment", "", "", "", ""],
  ["Overhead Work", "", "", "", ""],
  ["Excavation/Trenching", "", "", "", ""],
  ["Noise > 85 dBA", "", "", "", ""],
  ["Manual Lifting", "", "", "", ""],
  ["Weather Conditions", "", "", "", ""],
  ["Other:", "", "", "", ""],
  ["", "", "", "", ""],
  ["REQUIRED PPE", "YES/NO", "TOOLS & EQUIPMENT NEEDED", "YES/NO", ""],
  ["Hard Hat", "", "Ladders", "", ""],
  ["Safety Glasses", "", "Scaffolding", "", ""],
  ["Safety Vest", "", "Power Tools", "", ""],
  ["Safety Footwear", "", "Hand Tools", "", ""],
  ["Gloves", "", "Heavy Equipment", "", ""],
  ["Hearing Protection", "", "Lifts", "", ""],
  ["Respiratory Protection", "", "Fall Protection", "", ""],
  ["Face Shield", "", "Fire Extinguisher", "", ""],
  ["Other:", "", "Other:", "", ""],
  ["", "", "", "", ""],
  ["PERMITS REQUIRED", "YES/NO", "", "", ""],
  ["Hot Work", "", "", "", ""],
  ["Confined Space", "", "", "", ""],
  ["Excavation", "", "", "", ""],
  ["Lockout/Tagout", "", "", "", ""],
  ["Other:", "", "", "", ""],
  ["", "", "", "", ""],
  ["EMERGENCY RESPONSE PLAN", "", "", "", ""],
  ["Assembly Point:", "", "", "", ""],
  ["First Aid Location:", "", "", "", ""],
  ["Emergency Contact:", "", "", "", ""],
  ["", "", "", "", ""],
  ["Supervisor Signature:", "", "Date:", "", ""]
];

// Create the pre-task worksheet
const pretaskWS = XLSX.utils.aoa_to_sheet(pretaskData);

// Add column widths
pretaskWS['!cols'] = [
  { wch: 25 }, // A
  { wch: 15 }, // B
  { wch: 25 }, // C
  { wch: 15 }, // D
  { wch: 25 }  // E
];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, pretaskWS, "Pre-Task Plan");

// 4. INSPECTION LOG
const inspectionData = [
  ["SAFETY INSPECTION LOG", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["Project Name:", "", "Project #:", "", "", "", ""],
  ["Project Location:", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["DATE", "AREA INSPECTED", "INSPECTED BY", "FINDINGS/HAZARDS", "CORRECTIVE ACTIONS", "RESPONSIBLE PERSON", "COMPLETION DATE"],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["", "", "", "", "", "", ""],
  ["Project Manager Review:", "", "Date:", "", "", "", ""],
  ["Safety Manager Review:", "", "Date:", "", "", "", ""]
];

// Create the inspection log worksheet
const inspectionWS = XLSX.utils.aoa_to_sheet(inspectionData);

// Add column widths
inspectionWS['!cols'] = [
  { wch: 15 }, // A
  { wch: 20 }, // B
  { wch: 20 }, // C
  { wch: 30 }, // D
  { wch: 30 }, // E
  { wch: 20 }, // F
  { wch: 15 }  // G
];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, inspectionWS, "Inspection Log");

// 5. OBSERVATION REPORT
const observationData = [
  ["SAFETY OBSERVATION REPORT", "", "", "", ""],
  ["", "", "", "", ""],
  ["Project Name:", "", "Project #:", "", ""],
  ["Project Location:", "", "Date:", "", ""],
  ["Observer Name:", "", "Position:", "", ""],
  ["", "", "", "", ""],
  ["Type of Observation:", "", "", "", ""],
  ["☐ Planned   ☐ Unplanned", "", "", "", ""],
  ["Area/Location:", "", "", "", ""],
  ["Task Observed:", "", "", "", ""],
  ["", "", "", "", ""],
  ["SAFE PRACTICES OBSERVED", "", "", "", ""],
  ["", "", "", "", ""],
  ["1.", "", "", "", ""],
  ["2.", "", "", "", ""],
  ["3.", "", "", "", ""],
  ["4.", "", "", "", ""],
  ["", "", "", "", ""],
  ["AT-RISK BEHAVIORS/CONDITIONS OBSERVED", "CORRECTIVE ACTION TAKEN", "FOLLOW-UP REQUIRED", "RESPONSIBLE PERSON", "DUE DATE"],
  ["1.", "", "", "", ""],
  ["2.", "", "", "", ""],
  ["3.", "", "", "", ""],
  ["4.", "", "", "", ""],
  ["", "", "", "", ""],
  ["OBSERVATION CATEGORIES", "YES", "NO", "N/A", "COMMENTS"],
  ["1. PPE Used Properly", "", "", "", ""],
  ["2. Tools & Equipment in Good Condition", "", "", "", ""],
  ["3. Proper Body Positioning/Mechanics", "", "", "", ""],
  ["4. Work Area Clean & Orderly", "", "", "", ""],
  ["5. Procedures Being Followed", "", "", "", ""],
  ["6. Hazard Controls in Place", "", "", "", ""],
  ["7. Communication Between Workers", "", "", "", ""],
  ["8. Environmental Conditions Addressed", "", "", "", ""],
  ["", "", "", "", ""],
  ["Feedback Provided To:", "", "", "", ""],
  ["Feedback Summary:", "", "", "", ""],
  ["", "", "", "", ""],
  ["Observer Signature:", "", "Date:", "", ""],
  ["Supervisor Review:", "", "Date:", "", ""],
  ["Safety Manager Review:", "", "Date:", "", ""]
];

// Create the observation report worksheet
const observationWS = XLSX.utils.aoa_to_sheet(observationData);

// Add column widths
observationWS['!cols'] = [
  { wch: 30 }, // A
  { wch: 25 }, // B
  { wch: 20 }, // C
  { wch: 20 }, // D
  { wch: 15 }  // E
];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, observationWS, "Observation Report");

// Create a cover sheet with instructions
const coverData = [
  ["CONSTRUCTION SITE SAFETY PROGRAM", "", "", ""],
  ["", "", "", ""],
  ["Company Name:", "", "", ""],
  ["Project Name:", "", "", ""],
  ["Project Location:", "", "", ""],
  ["Project Manager:", "", "", ""],
  ["Safety Manager:", "", "", ""],
  ["Emergency Contact:", "", "", ""],
  ["", "", "", ""],
  ["WORKBOOK CONTENTS:", "", "", ""],
  ["", "", "", ""],
  ["1. Employee Orientation Form", "Document used to record safety orientation for new employees", "", ""],
  ["2. Job Hazard Analysis (JHA)", "Used to identify and control hazards for specific job tasks", "", ""],
  ["3. Pre-Task Plan", "Daily planning tool to identify hazards before starting work", "", ""],
  ["4. Inspection Log", "Record of safety inspections conducted on the jobsite", "", ""],
  ["5. Observation Report", "Form to document safety observations and feedback", "", ""],
  ["", "", "", ""],
  ["INSTRUCTIONS:", "", "", ""],
  ["", "", "", ""],
  ["1. Complete all applicable forms as required by your safety program.", "", "", ""],
  ["2. Maintain records of all completed forms for the duration of the project.", "", "", ""],
  ["3. Review safety documentation regularly to identify trends and areas for improvement.", "", "", ""],
  ["4. Share relevant safety information with all project personnel.", "", "", ""],
  ["5. Update forms as needed to address changing conditions or requirements.", "", "", ""],
  ["", "", "", ""],
  ["REMEMBER: SAFETY IS EVERYONE'S RESPONSIBILITY", "", "", ""]
];

// Create the cover sheet worksheet
const coverWS = XLSX.utils.aoa_to_sheet(coverData);

// Add column widths
coverWS['!cols'] = [
  { wch: 30 }, // A
  { wch: 40 }, // B
  { wch: 20 }, // C
  { wch: 20 }  // D
];

// Add the worksheet to the beginning of the workbook
XLSX.utils.book_append_sheet(workbook, coverWS, "Cover Sheet");

// Move the cover sheet to the front
const sheets = workbook.SheetNames;
const lastSheet = sheets.pop(); // Remove the last sheet (Cover Sheet)
sheets.unshift(lastSheet); // Add it to the beginning
workbook.SheetNames = sheets; // Update the sheet names

// Generate the Excel file
const excelOutput = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });

// Convert binary string to ArrayBuffer
function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) {
    view[i] = s.charCodeAt(i) & 0xFF;
  }
  return buf;
}

// Create a download link
const blob = new Blob([s2ab(excelOutput)], { type: 'application/octet-stream' });
const url = URL.createObjectURL(blob);

// This code would normally trigger a download, but in this environment, we'll return the URL
console.log("Excel workbook created successfully!");
// In a browser environment, you would use:
// const a = document.createElement('a');
// a.href = url;
// a.download = 'Construction_Safety_Program.xlsx';
// a.click();
