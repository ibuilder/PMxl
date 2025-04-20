// This script will create a complete Excel workbook for a construction project schedule
import * as XLSX from 'xlsx';

// Create a new workbook
const workbook = XLSX.utils.book_new();

// ===== TASK LIST SHEET =====
// Sample construction project tasks data 
const taskData = [
  ['ID', 'Task Name', 'Description', 'Responsible', 'Duration (days)', 'Start Date', 'End Date', 'Dependencies', 'Status', 'Priority', 'Notes'],
  ['1', 'Project Initiation', 'Initial project setup and paperwork', 'Project Manager', '10', '2025-05-01', '2025-05-14', '', 'Not Started', 'High', ''],
  ['2', 'Site Preparation', 'Clear the site and prepare for foundation work', 'Site Supervisor', '15', '2025-05-15', '2025-06-04', '1', 'Not Started', 'High', ''],
  ['3', 'Foundation Work', 'Excavation and foundation construction', 'Civil Engineer', '20', '2025-06-05', '2025-07-02', '2', 'Not Started', 'High', ''],
  ['4', 'Structural Framing', 'Construct main building frame', 'Structural Engineer', '30', '2025-07-03', '2025-08-13', '3', 'Not Started', 'High', ''],
  ['5', 'Roof Installation', 'Install roofing system', 'Roofing Contractor', '15', '2025-08-14', '2025-09-03', '4', 'Not Started', 'Medium', ''],
  ['6', 'Exterior Walls', 'Construct exterior walls and cladding', 'Construction Team', '20', '2025-08-14', '2025-09-10', '4', 'Not Started', 'Medium', ''],
  ['7', 'Plumbing Rough-In', 'Initial plumbing installation', 'Plumbing Contractor', '15', '2025-09-11', '2025-10-01', '4', 'Not Started', 'Medium', ''],
  ['8', 'Electrical Rough-In', 'Initial electrical systems installation', 'Electrical Contractor', '15', '2025-09-11', '2025-10-01', '4', 'Not Started', 'Medium', ''],
  ['9', 'HVAC Installation', 'Install heating and cooling systems', 'HVAC Contractor', '20', '2025-09-11', '2025-10-08', '4', 'Not Started', 'Medium', ''],
  ['10', 'Insulation', 'Install building insulation', 'Insulation Contractor', '10', '2025-10-09', '2025-10-22', '6,7,8,9', 'Not Started', 'Medium', ''],
  ['11', 'Drywall', 'Install interior wall panels', 'Drywall Contractor', '15', '2025-10-23', '2025-11-12', '10', 'Not Started', 'Medium', ''],
  ['12', 'Interior Finishes', 'Paint, trim, flooring and fixtures', 'Interior Contractor', '25', '2025-11-13', '2025-12-17', '11', 'Not Started', 'Medium', ''],
  ['13', 'Final Plumbing', 'Complete plumbing fixtures and connections', 'Plumbing Contractor', '10', '2025-12-18', '2025-12-31', '12', 'Not Started', 'Medium', ''],
  ['14', 'Final Electrical', 'Complete electrical fixtures and connections', 'Electrical Contractor', '10', '2025-12-18', '2025-12-31', '12', 'Not Started', 'Medium', ''],
  ['15', 'Site Cleanup', 'Final site cleanup and preparation', 'Site Supervisor', '5', '2026-01-01', '2026-01-07', '13,14', 'Not Started', 'Low', ''],
  ['16', 'Final Inspection', 'Regulatory inspections and approvals', 'Project Manager', '5', '2026-01-08', '2026-01-14', '15', 'Not Started', 'High', ''],
  ['17', 'Project Handover', 'Complete documentation and client handover', 'Project Manager', '3', '2026-01-15', '2026-01-19', '16', 'Not Started', 'High', '']
];

// Create task list worksheet
const taskSheet = XLSX.utils.aoa_to_sheet(taskData);

// Set column widths for better readability
const taskColWidths = [
  { wch: 5 },   // ID
  { wch: 20 },  // Task Name
  { wch: 35 },  // Description
  { wch: 20 },  // Responsible
  { wch: 15 },  // Duration
  { wch: 12 },  // Start Date
  { wch: 12 },  // End Date
  { wch: 15 },  // Dependencies
  { wch: 15 },  // Status
  { wch: 10 },  // Priority
  { wch: 25 }   // Notes
];
taskSheet['!cols'] = taskColWidths;

// Add task list to workbook
XLSX.utils.book_append_sheet(workbook, taskSheet, 'Task List');

// ===== GANTT CHART SHEET =====
// Create Gantt chart headers
const ganttData = [
  ['ID', 'Task Name', 'Start Date', 'End Date']
];

// Add data rows from task data (excluding header)
for (let i = 1; i < taskData.length; i++) {
  ganttData.push([
    taskData[i][0],  // ID
    taskData[i][1],  // Task Name
    taskData[i][5],  // Start Date
    taskData[i][6]   // End Date
  ]);
}

// Create Gantt chart worksheet
const ganttSheet = XLSX.utils.aoa_to_sheet(ganttData);

// Set column widths
const ganttColWidths = [
  { wch: 5 },   // ID
  { wch: 20 },  // Task Name
  { wch: 12 },  // Start Date
  { wch: 12 }   // End Date
];
ganttSheet['!cols'] = ganttColWidths;

// Add Gantt chart sheet to workbook
XLSX.utils.book_append_sheet(workbook, ganttSheet, 'Gantt Chart');

// ===== MILESTONES SHEET =====
// Create milestone data
const milestoneData = [
  ['ID', 'Milestone Name', 'Target Date', 'Responsible', 'Associated Tasks', 'Status', 'Notes'],
  ['M1', 'Project Start', '2025-05-01', 'Project Manager', '1', 'Not Started', ''],
  ['M2', 'Foundation Complete', '2025-07-02', 'Civil Engineer', '3', 'Not Started', ''],
  ['M3', 'Structure Complete', '2025-08-13', 'Structural Engineer', '4', 'Not Started', ''],
  ['M4', 'Building Enclosed', '2025-09-10', 'Construction Team', '5,6', 'Not Started', ''],
  ['M5', 'MEP Rough-in Complete', '2025-10-08', 'MEP Coordinator', '7,8,9', 'Not Started', ''],
  ['M6', 'Interior Finishes Complete', '2025-12-17', 'Interior Contractor', '12', 'Not Started', ''],
  ['M7', 'Final Inspections', '2026-01-14', 'Project Manager', '16', 'Not Started', ''],
  ['M8', 'Project Completion', '2026-01-19', 'Project Manager', '17', 'Not Started', '']
];

// Create milestone worksheet
const milestoneSheet = XLSX.utils.aoa_to_sheet(milestoneData);

// Set column widths
const milestoneColWidths = [
  { wch: 5 },   // ID
  { wch: 25 },  // Milestone Name
  { wch: 12 },  // Target Date
  { wch: 20 },  // Responsible
  { wch: 15 },  // Associated Tasks
  { wch: 15 },  // Status
  { wch: 25 }   // Notes
];
milestoneSheet['!cols'] = milestoneColWidths;

// Add milestones sheet to workbook
XLSX.utils.book_append_sheet(workbook, milestoneSheet, 'Milestones');

// ===== CONSTRAINTS SHEET =====
// Create constraints data
const constraintData = [
  ['ID', 'Constraint Type', 'Description', 'Impact', 'Affected Tasks', 'Mitigation Plan', 'Status', 'Responsible'],
  ['C1', 'Weather', 'Winter conditions may delay exterior work', 'Schedule delay', '5,6', 'Schedule exterior work in warmer months, prepare contingency plan', 'Active', 'Project Manager'],
  ['C2', 'Regulatory', 'Building permit approval process', 'Cannot start construction without permits', '2,3', 'Submit applications early, follow up regularly', 'Active', 'Permit Coordinator'],
  ['C3', 'Resource', 'Limited skilled labor availability', 'Potential delays in specialized work', '7,8,9', 'Pre-book contractors, consider alternative sourcing', 'Active', 'Resource Manager'],
  ['C4', 'Budget', 'Material cost fluctuations', 'Budget overruns', 'All', 'Secure price commitments early, include contingency budget', 'Active', 'Financial Manager'],
  ['C5', 'Site', 'Limited site access for deliveries', 'Logistics complications', 'All', 'Create detailed delivery schedule, coordinate with neighbors', 'Active', 'Site Supervisor'],
  ['C6', 'Technical', 'Complex foundation requirements', 'Additional engineering required', '3', 'Early engagement with geotechnical experts', 'Active', 'Civil Engineer'],
  ['C7', 'Environmental', 'Noise restrictions during certain hours', 'Limited working hours', 'All', 'Schedule noisy work during permitted hours, notify community', 'Active', 'Site Supervisor']
];

// Create constraints worksheet
const constraintSheet = XLSX.utils.aoa_to_sheet(constraintData);

// Set column widths
const constraintColWidths = [
  { wch: 5 },   // ID
  { wch: 15 },  // Constraint Type
  { wch: 35 },  // Description
  { wch: 25 },  // Impact
  { wch: 15 },  // Affected Tasks
  { wch: 35 },  // Mitigation Plan
  { wch: 10 },  // Status
  { wch: 20 }   // Responsible
];
constraintSheet['!cols'] = constraintColWidths;

// Add constraints sheet to workbook
XLSX.utils.book_append_sheet(workbook, constraintSheet, 'Constraints');

// ===== LOGISTIC PHASING SHEET =====
// Create logistics phasing data
const phasingData = [
  ['Phase', 'Start Date', 'End Date', 'Description', 'Tasks Involved', 'Resources Required', 'Site Access Points', 'Storage Areas', 'Special Requirements'],
  ['Phase 1', '2025-05-01', '2025-07-02', 'Site Preparation & Foundation', '1,2,3', 'Excavators, Concrete trucks, Laborers', 'Main entrance', 'North corner of site', 'Temporary fencing, Silt control'],
  ['Phase 2', '2025-07-03', '2025-09-10', 'Structure & Envelope', '4,5,6', 'Crane, Delivery trucks, Framing team', 'Main and east entrances', 'East side of building', 'Staging area for materials, Tower crane setup'],
  ['Phase 3', '2025-09-11', '2025-10-22', 'MEP Rough-In & Insulation', '7,8,9,10', 'Specialized contractors, Material deliveries', 'All entrances', 'Interior of building', 'Secure storage for equipment, Temporary power'],
  ['Phase 4', '2025-10-23', '2025-12-31', 'Interior Finishes & MEP Completion', '11,12,13,14', 'Finish contractors, Fixtures deliveries', 'Main entrance only', 'Interior secured rooms', 'Climate control active, Dust control measures'],
  ['Phase 5', '2026-01-01', '2026-01-19', 'Completion & Handover', '15,16,17', 'Cleaning crews, Inspection teams', 'Restricted access', 'Minimal on-site storage', 'Furniture deliveries, Systems testing']
];

// Create logistics phasing worksheet
const phasingSheet = XLSX.utils.aoa_to_sheet(phasingData);

// Set column widths
const phasingColWidths = [
  { wch: 10 },  // Phase
  { wch: 12 },  // Start Date
  { wch: 12 },  // End Date
  { wch: 25 },  // Description
  { wch: 15 },  // Tasks Involved
  { wch: 30 },  // Resources Required
  { wch: 20 },  // Site Access Points
  { wch: 20 },  // Storage Areas
  { wch: 35 }   // Special Requirements
];
phasingSheet['!cols'] = phasingColWidths;

// Add logistics phasing sheet to workbook
XLSX.utils.book_append_sheet(workbook, phasingSheet, 'Logistic Phasing');

// Convert workbook to binary Excel format
const excelData = XLSX.write(workbook, { type: 'binary', bookType: 'xlsx' });

// Output: the workbook object has been created with all required sheets
// In a real application, you would save this to a file
console.log("Construction Project Schedule workbook created successfully with the following sheets:");
console.log("- Task List: Detailed list of all project tasks");
console.log("- Gantt Chart: Visual timeline of tasks");
console.log("- Milestones: Key project milestones");
console.log("- Constraints: Project constraints and mitigation plans");
console.log("- Logistic Phasing: Construction phases with logistics details");

// Note: This code generates the workbook structure
// To use in Excel, you would need to convert this to a downloadable file
// or add visualization for the Gantt chart
