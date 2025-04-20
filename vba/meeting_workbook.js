import * as XLSX from 'xlsx';

// Create a new workbook
const workbook = XLSX.utils.book_new();

// Meeting Log Sheet
const meetingLogData = [
  ['Meeting Log', '', '', '', '', ''],
  ['', '', '', '', '', ''],
  ['Meeting ID', 'Date', 'Time', 'Location', 'Attendees', 'Purpose'],
  ['ML001', '2025-04-20', '09:00-10:00', 'Conference Room A', 'Team Members', 'Project Kickoff'],
  ['ML002', '2025-04-27', '14:00-15:00', 'Virtual Meeting', 'Department Heads', 'Monthly Review'],
  ['', '', '', '', '', ''],
  ['', '', '', '', '', ''],
  ['', '', '', '', '', ''],
  ['', '', '', '', '', ''],
  ['', '', '', '', '', '']
];

// Meeting Agenda Sheet
const meetingAgendaData = [
  ['Meeting Agenda', '', '', '', ''],
  ['', '', '', '', ''],
  ['Meeting Title:', '', '', '', ''],
  ['Date:', '', '', '', ''],
  ['Time:', '', '', '', ''],
  ['Location:', '', '', '', ''],
  ['Meeting Called By:', '', '', '', ''],
  ['Attendees:', '', '', '', ''],
  ['', '', '', '', ''],
  ['Objective:', '', '', '', ''],
  ['', '', '', '', ''],
  ['Agenda Items:', '', 'Duration', 'Presenter', 'Notes'],
  ['1.', '', '', '', ''],
  ['2.', '', '', '', ''],
  ['3.', '', '', '', ''],
  ['4.', '', '', '', ''],
  ['5.', '', '', '', ''],
  ['', '', '', '', ''],
  ['Pre-meeting Preparation:', '', '', '', ''],
  ['', '', '', '', ''],
  ['Additional Information:', '', '', '', '']
];

// Meeting Minutes Sheet
const meetingMinutesData = [
  ['Meeting Minutes', '', '', '', ''],
  ['', '', '', '', ''],
  ['Meeting Title:', '', '', '', ''],
  ['Date:', '', '', '', ''],
  ['Time:', '', '', '', ''],
  ['Location:', '', '', '', ''],
  ['Attendees Present:', '', '', '', ''],
  ['Attendees Absent:', '', '', '', ''],
  ['', '', '', '', ''],
  ['Agenda Items', 'Discussion Points', 'Action Items', 'Person Responsible', 'Deadline'],
  ['1.', '', '', '', ''],
  ['', '', '', '', ''],
  ['2.', '', '', '', ''],
  ['', '', '', '', ''],
  ['3.', '', '', '', ''],
  ['', '', '', '', ''],
  ['4.', '', '', '', ''],
  ['', '', '', '', ''],
  ['5.', '', '', '', ''],
  ['', '', '', '', ''],
  ['', '', '', '', ''],
  ['Next Meeting:', '', '', '', ''],
  ['Date:', '', '', '', ''],
  ['Time:', '', '', '', ''],
  ['Location:', '', '', '', ''],
  ['Agenda Items:', '', '', '', '']
];

// Create worksheets from the data
const meetingLogSheet = XLSX.utils.aoa_to_sheet(meetingLogData);
const meetingAgendaSheet = XLSX.utils.aoa_to_sheet(meetingAgendaData);
const meetingMinutesSheet = XLSX.utils.aoa_to_sheet(meetingMinutesData);

// Add formatting
// Meeting Log formatting
meetingLogSheet['!merges'] = [
  {s: {r: 0, c: 0}, e: {r: 0, c: 5}}, // Merge title row across all columns
];
meetingLogSheet['!cols'] = [
  {wch: 15}, // Column A width
  {wch: 15}, // Column B width
  {wch: 20}, // Column C width
  {wch: 20}, // Column D width
  {wch: 25}, // Column E width
  {wch: 25}, // Column F width
];
meetingLogSheet['!rows'] = [{hpt: 30}]; // Header row height

// Meeting Agenda formatting
meetingAgendaSheet['!merges'] = [
  {s: {r: 0, c: 0}, e: {r: 0, c: 4}}, // Merge title row across all columns
  {s: {r: 3, c: 1}, e: {r: 3, c: 4}}, // Merge Date field
  {s: {r: 4, c: 1}, e: {r: 4, c: 4}}, // Merge Time field
  {s: {r: 5, c: 1}, e: {r: 5, c: 4}}, // Merge Location field
  {s: {r: 6, c: 1}, e: {r: 6, c: 4}}, // Merge Called By field
  {s: {r: 7, c: 1}, e: {r: 7, c: 4}}, // Merge Attendees field
  {s: {r: 9, c: 1}, e: {r: 9, c: 4}}, // Merge Objective field
];
meetingAgendaSheet['!cols'] = [
  {wch: 20}, // Column A width
  {wch: 30}, // Column B width
  {wch: 15}, // Column C width
  {wch: 15}, // Column D width
  {wch: 30}, // Column E width
];

// Meeting Minutes formatting
meetingMinutesSheet['!merges'] = [
  {s: {r: 0, c: 0}, e: {r: 0, c: 4}}, // Merge title row across all columns
  {s: {r: 3, c: 1}, e: {r: 3, c: 4}}, // Merge Date field
  {s: {r: 4, c: 1}, e: {r: 4, c: 4}}, // Merge Time field
  {s: {r: 5, c: 1}, e: {r: 5, c: 4}}, // Merge Location field
  {s: {r: 6, c: 1}, e: {r: 6, c: 4}}, // Merge Attendees Present field
  {s: {r: 7, c: 1}, e: {r: 7, c: 4}}, // Merge Attendees Absent field
];
meetingMinutesSheet['!cols'] = [
  {wch: 20}, // Column A width
  {wch: 30}, // Column B width
  {wch: 20}, // Column C width
  {wch: 20}, // Column D width
  {wch: 15}, // Column E width
];

// Add the worksheets to the workbook
XLSX.utils.book_append_sheet(workbook, meetingLogSheet, 'Meeting Log');
XLSX.utils.book_append_sheet(workbook, meetingAgendaSheet, 'Meeting Agenda');
XLSX.utils.book_append_sheet(workbook, meetingMinutesSheet, 'Meeting Minutes');

// Write the workbook to a file
const workbookOutput = XLSX.write(workbook, { type: 'binary', bookType: 'xlsx' });

// Convert binary to blob data for download
const buffer = new ArrayBuffer(workbookOutput.length);
const view = new Uint8Array(buffer);
for (let i = 0; i < workbookOutput.length; i++) {
  view[i] = workbookOutput.charCodeAt(i) & 0xFF;
}

// Create a Blob from the ArrayBuffer
const blob = new Blob([buffer], { type: 'application/octet-stream' });

// Create a link for downloading
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
document.body.appendChild(a);
a.style = 'display: none';
a.href = url;
a.download = 'Meeting_Management_Workbook.xlsx';
a.click();

// Clean up
URL.revokeObjectURL(url);
document.body.removeChild(a);
