// This JavaScript code generates an Excel workbook for construction budget management
// You can copy and paste this into Node.js with the ExcelJS library installed

const ExcelJS = require('exceljs');

async function createConstructionBudgetWorkbook() {
  // Create a new workbook
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Construction Budget Tool';
  workbook.lastModifiedBy = 'Project Manager';
  workbook.created = new Date();
  workbook.modified = new Date();

  // Create the Budget Plan sheet
  const budgetSheet = workbook.addWorksheet('Budget Plan', {
    properties: { tabColor: { argb: '4F81BD' } }
  });

  // Set column widths
  budgetSheet.columns = [
    { header: 'Category', key: 'category', width: 20 },
    { header: 'Subcategory', key: 'subcategory', width: 25 },
    { header: 'Description', key: 'description', width: 30 },
    { header: 'Estimated Cost', key: 'estimated', width: 15 },
    { header: 'Unit', key: 'unit', width: 10 },
    { header: 'Quantity', key: 'quantity', width: 10 },
    { header: 'Unit Price', key: 'unitprice', width: 12 },
    { header: 'Total Budget', key: 'total', width: 15 },
    { header: 'Notes', key: 'notes', width: 30 }
  ];

  // Style the header row
  budgetSheet.getRow(1).font = { bold: true, size: 12 };
  budgetSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '4F81BD' }
  };
  budgetSheet.getRow(1).font = { color: { argb: 'FFFFFF' }, bold: true };

  // Add sample data categories for construction budget
  const budgetCategories = [
    { category: 'Pre-Construction', subcategory: 'Site Survey', description: 'Topographical survey of site', estimated: '=G2*H2', unit: 'Acres', quantity: 5, unitprice: 1000, total: '=D2', notes: '' },
    { category: 'Pre-Construction', subcategory: 'Permits', description: 'Building permits', estimated: '=G3*H3', unit: 'Lump Sum', quantity: 1, unitprice: 15000, total: '=D3', notes: 'May vary by municipality' },
    { category: 'Pre-Construction', subcategory: 'Architecture/Engineering', description: 'Design services', estimated: '=G4*H4', unit: 'Lump Sum', quantity: 1, unitprice: 75000, total: '=D4', notes: '' },
    { category: 'Site Work', subcategory: 'Excavation', description: 'Site preparation and grading', estimated: '=G5*H5', unit: 'Cu. Yards', quantity: 500, unitprice: 30, total: '=D5', notes: 'Includes equipment rental' },
    { category: 'Site Work', subcategory: 'Utilities', description: 'Water, sewer, electric connections', estimated: '=G6*H6', unit: 'Lump Sum', quantity: 1, unitprice: 35000, total: '=D6', notes: '' },
    { category: 'Foundation', subcategory: 'Concrete', description: 'Foundation concrete', estimated: '=G7*H7', unit: 'Cu. Yards', quantity: 120, unitprice: 200, total: '=D7', notes: 'Includes labor' },
    { category: 'Foundation', subcategory: 'Waterproofing', description: 'Foundation waterproofing', estimated: '=G8*H8', unit: 'Sq. Ft.', quantity: 2000, unitprice: 5, total: '=D8', notes: '' },
    { category: 'Framing', subcategory: 'Wood Framing', description: 'Structural framing materials', estimated: '=G9*H9', unit: 'Sq. Ft.', quantity: 3500, unitprice: 25, total: '=D9', notes: '' },
    { category: 'Framing', subcategory: 'Labor', description: 'Framing labor', estimated: '=G10*H10', unit: 'Hours', quantity: 500, unitprice: 50, total: '=D10', notes: '' },
    { category: 'Exterior', subcategory: 'Roofing', description: 'Roofing materials and installation', estimated: '=G11*H11', unit: 'Sq. Ft.', quantity: 2000, unitprice: 15, total: '=D11', notes: '' },
    { category: 'Exterior', subcategory: 'Siding', description: 'Exterior siding', estimated: '=G12*H12', unit: 'Sq. Ft.', quantity: 3000, unitprice: 12, total: '=D12', notes: '' },
    { category: 'Exterior', subcategory: 'Windows & Doors', description: 'Exterior doors and windows', estimated: '=G13*H13', unit: 'Each', quantity: 25, unitprice: 1200, total: '=D13', notes: '' },
    { category: 'Interior', subcategory: 'Drywall', description: 'Drywall installation', estimated: '=G14*H14', unit: 'Sq. Ft.', quantity: 5000, unitprice: 4, total: '=D14', notes: '' },
    { category: 'Interior', subcategory: 'Flooring', description: 'Various flooring materials', estimated: '=G15*H15', unit: 'Sq. Ft.', quantity: 3500, unitprice: 12, total: '=D15', notes: '' },
    { category: 'Interior', subcategory: 'Painting', description: 'Interior painting', estimated: '=G16*H16', unit: 'Sq. Ft.', quantity: 5000, unitprice: 3, total: '=D16', notes: '' },
    { category: 'Interior', subcategory: 'Trim & Doors', description: 'Interior trim and doors', estimated: '=G17*H17', unit: 'Lump Sum', quantity: 1, unitprice: 22000, total: '=D17', notes: '' },
    { category: 'MEP', subcategory: 'Electrical', description: 'Electrical systems', estimated: '=G18*H18', unit: 'Lump Sum', quantity: 1, unitprice: 45000, total: '=D18', notes: '' },
    { category: 'MEP', subcategory: 'Plumbing', description: 'Plumbing systems', estimated: '=G19*H19', unit: 'Lump Sum', quantity: 1, unitprice: 38000, total: '=D19', notes: '' },
    { category: 'MEP', subcategory: 'HVAC', description: 'Heating and cooling systems', estimated: '=G20*H20', unit: 'Lump Sum', quantity: 1, unitprice: 42000, total: '=D20', notes: '' },
    { category: 'Fixtures', subcategory: 'Kitchen', description: 'Kitchen cabinets and countertops', estimated: '=G21*H21', unit: 'Lump Sum', quantity: 1, unitprice: 30000, total: '=D21', notes: '' },
    { category: 'Fixtures', subcategory: 'Bathroom', description: 'Bathroom fixtures and tile', estimated: '=G22*H22', unit: 'Each', quantity: 3, unitprice: 12000, total: '=D22', notes: '' },
    { category: 'Fixtures', subcategory: 'Appliances', description: 'Kitchen appliances', estimated: '=G23*H23', unit: 'Lump Sum', quantity: 1, unitprice: 15000, total: '=D23', notes: '' },
    { category: 'Landscaping', subcategory: 'Grading', description: 'Final grading', estimated: '=G24*H24', unit: 'Sq. Ft.', quantity: 10000, unitprice: 0.75, total: '=D24', notes: '' },
    { category: 'Landscaping', subcategory: 'Planting', description: 'Trees, shrubs, and lawn', estimated: '=G25*H25', unit: 'Lump Sum', quantity: 1, unitprice: 15000, total: '=D25', notes: '' },
    { category: 'Other', subcategory: 'Cleaning', description: 'Final cleaning', estimated: '=G26*H26', unit: 'Lump Sum', quantity: 1, unitprice: 3500, total: '=D26', notes: '' },
    { category: 'Other', subcategory: 'Dumpsters', description: 'Waste removal', estimated: '=G27*H27', unit: 'Each', quantity: 10, unitprice: 600, total: '=D27', notes: '' },
    { category: 'Contingency', subcategory: 'Contingency Fund', description: '10% of total budget', estimated: '=G28*H28', unit: 'Percentage', quantity: 0.1, unitprice: '=SUM(H2:H27)', total: '=D28', notes: 'For unexpected costs' }
  ];

  // Add the sample data to the budget sheet
  budgetSheet.addRows(budgetCategories);

  // Add a total row at the bottom
  const totalRow = budgetSheet.addRow(['TOTAL', '', '', '', '', '', '', '=SUM(H2:H28)', '']);
  totalRow.font = { bold: true };
  totalRow.getCell(8).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'E6E6E6' }
  };

  // Apply number formatting
  for (let i = 2; i <= budgetSheet.rowCount; i++) {
    budgetSheet.getCell(`D${i}`).numFmt = '$#,##0.00';
    budgetSheet.getCell(`H${i}`).numFmt = '$#,##0.00';
    budgetSheet.getCell(`G${i}`).numFmt = '#,##0.00';
  }

  // Format with alternating row colors
  for (let i = 2; i <= budgetSheet.rowCount; i++) {
    if (i % 2 === 0) {
      budgetSheet.getRow(i).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' }
      };
    }
  }

  // Add borders to all cells
  for (let i = 1; i <= budgetSheet.rowCount; i++) {
    for (let j = 1; j <= 9; j++) {
      budgetSheet.getRow(i).getCell(j).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
  }

  // Create Budget Forecast Sheet
  const forecastSheet = workbook.addWorksheet('Budget Forecast', {
    properties: { tabColor: { argb: '9BBB59' } }
  });

  // Set up columns for the forecast sheet
  forecastSheet.columns = [
    { header: 'Category', key: 'category', width: 20 },
    { header: 'Subcategory', key: 'subcategory', width: 25 },
    { header: 'Total Budget', key: 'budget', width: 15 },
    { header: 'Month 1', key: 'month1', width: 12 },
    { header: 'Month 2', key: 'month2', width: 12 },
    { header: 'Month 3', key: 'month3', width: 12 },
    { header: 'Month 4', key: 'month4', width: 12 },
    { header: 'Month 5', key: 'month5', width: 12 },
    { header: 'Month 6', key: 'month6', width: 12 },
    { header: 'Month 7', key: 'month7', width: 12 },
    { header: 'Month 8', key: 'month8', width: 12 },
    { header: 'Total Forecast', key: 'totalforecast', width: 15 },
    { header: 'Variance', key: 'variance', width: 15 }
  ];

  // Style the header row
  forecastSheet.getRow(1).font = { bold: true, size: 12 };
  forecastSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '9BBB59' }
  };
  forecastSheet.getRow(1).font = { color: { argb: 'FFFFFF' }, bold: true };

  // Create data with references to Budget Plan sheet
  const forecastData = [];
  for (let i = 2; i <= 28; i++) {
    forecastData.push({
      category: `='Budget Plan'!A${i}`,
      subcategory: `='Budget Plan'!B${i}`,
      budget: `='Budget Plan'!H${i}`,
      month1: 0,
      month2: 0,
      month3: 0,
      month4: 0,
      month5: 0,
      month6: 0,
      month7: 0,
      month8: 0,
      totalforecast: `=SUM(D${i}:J${i})`,
      variance: `=C${i}-K${i}`
    });
  }

  // Add the forecast data
  forecastSheet.addRows(forecastData);

  // Distribute budget across months (sample data)
  // Pre-Construction - mostly in month 1
  forecastSheet.getCell('D2').value = '=C2 * 1.0'; // Site Survey - Month 1 (100%)
  forecastSheet.getCell('D3').value = '=C3 * 1.0'; // Permits - Month 1 (100%)
  forecastSheet.getCell('D4').value = '=C4 * 0.6'; // Architecture - Month 1 (60%)
  forecastSheet.getCell('E4').value = '=C4 * 0.4'; // Architecture - Month 2 (40%)
  
  // Site Work - month 1-2
  forecastSheet.getCell('D5').value = '=C5 * 0.7'; // Excavation - Month 1 (70%)
  forecastSheet.getCell('E5').value = '=C5 * 0.3'; // Excavation - Month 2 (30%)
  forecastSheet.getCell('D6').value = '=C6 * 0.3'; // Utilities - Month 1 (30%)
  forecastSheet.getCell('E6').value = '=C6 * 0.7'; // Utilities - Month 2 (70%)
  
  // Foundation - month 2-3
  forecastSheet.getCell('E7').value = '=C7 * 0.6'; // Concrete - Month 2 (60%)
  forecastSheet.getCell('F7').value = '=C7 * 0.4'; // Concrete - Month 3 (40%)
  forecastSheet.getCell('E8').value = '=C8 * 0.5'; // Waterproofing - Month 2 (50%)
  forecastSheet.getCell('F8').value = '=C8 * 0.5'; // Waterproofing - Month 3 (50%)
  
  // Framing - month 3-4
  forecastSheet.getCell('F9').value = '=C9 * 0.6'; // Wood Framing - Month 3 (60%)
  forecastSheet.getCell('G9').value = '=C9 * 0.4'; // Wood Framing - Month 4 (40%)
  forecastSheet.getCell('F10').value = '=C10 * 0.6'; // Framing Labor - Month 3 (60%)
  forecastSheet.getCell('G10').value = '=C10 * 0.4'; // Framing Labor - Month 4 (40%)
  
  // Exterior - month 4-5
  forecastSheet.getCell('G11').value = '=C11 * 0.8'; // Roofing - Month 4 (80%)
  forecastSheet.getCell('H11').value = '=C11 * 0.2'; // Roofing - Month 5 (20%)
  forecastSheet.getCell('G12').value = '=C12 * 0.3'; // Siding - Month 4 (30%)
  forecastSheet.getCell('H12').value = '=C12 * 0.7'; // Siding - Month 5 (70%)
  forecastSheet.getCell('G13').value = '=C13 * 0.4'; // Windows & Doors - Month 4 (40%)
  forecastSheet.getCell('H13').value = '=C13 * 0.6'; // Windows & Doors - Month 5 (60%)
  
  // Interior - month 5-7
  forecastSheet.getCell('H14').value = '=C14 * 0.7'; // Drywall - Month 5 (70%)
  forecastSheet.getCell('I14').value = '=C14 * 0.3'; // Drywall - Month 6 (30%)
  forecastSheet.getCell('H15').value = '=C15 * 0.2'; // Flooring - Month 5 (20%)
  forecastSheet.getCell('I15').value = '=C15 * 0.6'; // Flooring - Month 6 (60%)
  forecastSheet.getCell('J15').value = '=C15 * 0.2'; // Flooring - Month 7 (20%)
  forecastSheet.getCell('I16').value = '=C16 * 0.7'; // Painting - Month 6 (70%)
  forecastSheet.getCell('J16').value = '=C16 * 0.3'; // Painting - Month 7 (30%)
  forecastSheet.getCell('I17').value = '=C17 * 0.6'; // Trim & Doors - Month 6 (60%)
  forecastSheet.getCell('J17').value = '=C17 * 0.4'; // Trim & Doors - Month 7 (40%)
  
  // MEP - month 3-6
  forecastSheet.getCell('F18').value = '=C18 * 0.2'; // Electrical - Month 3 (20%)
  forecastSheet.getCell('G18').value = '=C18 * 0.3'; // Electrical - Month 4 (30%)
  forecastSheet.getCell('H18').value = '=C18 * 0.3'; // Electrical - Month 5 (30%)
  forecastSheet.getCell('I18').value = '=C18 * 0.2'; // Electrical - Month 6 (20%)
  forecastSheet.getCell('F19').value = '=C19 * 0.2'; // Plumbing - Month 3 (20%)
  forecastSheet.getCell('G19').value = '=C19 * 0.3'; // Plumbing - Month 4 (30%)
  forecastSheet.getCell('H19').value = '=C19 * 0.3'; // Plumbing - Month 5 (30%)
  forecastSheet.getCell('I19').value = '=C19 * 0.2'; // Plumbing - Month 6 (20%)
  forecastSheet.getCell('G20').value = '=C20 * 0.3'; // HVAC - Month 4 (30%)
  forecastSheet.getCell('H20').value = '=C20 * 0.4'; // HVAC - Month 5 (40%)
  forecastSheet.getCell('I20').value = '=C20 * 0.3'; // HVAC - Month 6 (30%)
  
  // Fixtures - month 6-7
  forecastSheet.getCell('I21').value = '=C21 * 0.7'; // Kitchen - Month 6 (70%)
  forecastSheet.getCell('J21').value = '=C21 * 0.3'; // Kitchen - Month 7 (30%)
  forecastSheet.getCell('I22').value = '=C22 * 0.5'; // Bathroom - Month 6 (50%)
  forecastSheet.getCell('J22').value = '=C22 * 0.5'; // Bathroom - Month 7 (50%)
  forecastSheet.getCell('J23').value = '=C23 * 1.0'; // Appliances - Month 7 (100%)
  
  // Landscaping and other - month 7-8
  forecastSheet.getCell('J24').value = '=C24 * 0.3'; // Grading - Month 7 (30%)
  forecastSheet.getCell('K24').value = '=C24 * 0.7'; // Grading - Month 8 (70%)
  forecastSheet.getCell('K25').value = '=C25 * 1.0'; // Planting - Month 8 (100%)
  forecastSheet.getCell('K26').value = '=C26 * 1.0'; // Cleaning - Month 8 (100%)
  
  // Dumpsters - spread across project
  forecastSheet.getCell('D27').value = '=C27 * 0.1'; // Dumpsters - Month 1 (10%)
  forecastSheet.getCell('E27').value = '=C27 * 0.1'; // Dumpsters - Month 2 (10%)
  forecastSheet.getCell('F27').value = '=C27 * 0.1'; // Dumpsters - Month 3 (10%)
  forecastSheet.getCell('G27').value = '=C27 * 0.15'; // Dumpsters - Month 4 (15%)
  forecastSheet.getCell('H27').value = '=C27 * 0.15'; // Dumpsters - Month 5 (15%)
  forecastSheet.getCell('I27').value = '=C27 * 0.15'; // Dumpsters - Month 6 (15%)
  forecastSheet.getCell('J27').value = '=C27 * 0.15'; // Dumpsters - Month 7 (15%)
  forecastSheet.getCell('K27').value = '=C27 * 0.1'; // Dumpsters - Month 8 (10%)
  
  // Contingency - distribute proportionally
  forecastSheet.getCell('D28').value = '=SUM(D2:D27)/SUM(C2:C27)*C28'; // Contingency - Month 1
  forecastSheet.getCell('E28').value = '=SUM(E2:E27)/SUM(C2:C27)*C28'; // Contingency - Month 2
  forecastSheet.getCell('F28').value = '=SUM(F2:F27)/SUM(C2:C27)*C28'; // Contingency - Month 3
  forecastSheet.getCell('G28').value = '=SUM(G2:G27)/SUM(C2:C27)*C28'; // Contingency - Month 4
  forecastSheet.getCell('H28').value = '=SUM(H2:H27)/SUM(C2:C27)*C28'; // Contingency - Month 5
  forecastSheet.getCell('I28').value = '=SUM(I2:I27)/SUM(C2:C27)*C28'; // Contingency - Month 6
  forecastSheet.getCell('J28').value = '=SUM(J2:J27)/SUM(C2:C27)*C28'; // Contingency - Month 7
  forecastSheet.getCell('K28').value = '=SUM(K2:K27)/SUM(C2:C27)*C28'; // Contingency - Month 8

  // Add monthly totals row
  const forecastTotalRow = forecastSheet.addRow(['MONTHLY TOTALS', '', '=SUM(C2:C28)', '=SUM(D2:D28)', '=SUM(E2:E28)', '=SUM(F2:F28)', '=SUM(G2:G28)', '=SUM(H2:H28)', '=SUM(I2:I28)', '=SUM(J2:J28)', '=SUM(K2:K28)', '=SUM(L2:L28)', '=SUM(M2:M28)']);
  forecastTotalRow.font = { bold: true };
  
  // Add cumulative totals row
  const cumulativeTotalRow = forecastSheet.addRow(['CUMULATIVE TOTALS', '', '', '=D29', '=D30+E29', '=E30+F29', '=F30+G29', '=G30+H29', '=H30+I29', '=I30+J29', '=J30+K29', '', '']);
  cumulativeTotalRow.font = { bold: true };
  
  // Add percentage complete row
  const percentCompleteRow = forecastSheet.addRow(['PERCENTAGE COMPLETE', '', '', '=D30/C29', '=E30/C29', '=F30/C29', '=G30/C29', '=H30/C29', '=I30/C29', '=J30/C29', '=K30/C29', '', '']);
  percentCompleteRow.font = { bold: true };

  // Apply number formatting to forecast sheet
  for (let i = 2; i <= forecastSheet.rowCount; i++) {
    for (let j = 3; j <= 13; j++) {
      if (i <= 31) {
        forecastSheet.getRow(i).getCell(j).numFmt = '$#,##0.00';
      } else if (i === 31) {
        // Percentage row
        forecastSheet.getRow(i).getCell(j).numFmt = '0.00%';
      }
    }
  }

  // Format with alternating row colors
  for (let i = 2; i <= forecastSheet.rowCount; i++) {
    if (i % 2 === 0) {
      forecastSheet.getRow(i).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' }
      };
    }
  }
  
  // Add borders
  for (let i = 1; i <= forecastSheet.rowCount; i++) {
    for (let j = 1; j <= 13; j++) {
      forecastSheet.getRow(i).getCell(j).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
  }

  // Create Cost Tracking Sheet
  const costTrackingSheet = workbook.addWorksheet('Cost Tracking', {
    properties: { tabColor: { argb: 'C0504D' } }
  });

  // Set up columns for the cost tracking sheet
  costTrackingSheet.columns = [
    { header: 'Category', key: 'category', width: 20 },
    { header: 'Subcategory', key: 'subcategory', width: 25 },
    { header: 'Total Budget', key: 'budget', width: 15 },
    { header: 'Committed Cost', key: 'committed', width: 15 },
    { header: 'Invoiced to Date', key: 'invoiced', width: 15 },
    { header: 'Paid to Date', key: 'paid', width: 15 },
    { header: 'Remaining to Invoice', key: 'remaining', width: 18 },
    { header: 'Budget Variance', key: 'variance', width: 15 },
    { header: 'Percent Complete', key: 'complete', width: 15 },
    { header: 'Notes', key: 'notes', width: 30 }
  ];

  // Style the header row
  costTrackingSheet.getRow(1).font = { bold: true, size: 12 };
  costTrackingSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'C0504D' }
  };
  costTrackingSheet.getRow(1).font = { color: { argb: 'FFFFFF' }, bold: true };

  // Create data for cost tracking with references to Budget Plan
  const costTrackingData = [];
  for (let i = 2; i <= 28; i++) {
    costTrackingData.push({
      category: `='Budget Plan'!A${i}`,
      subcategory: `='Budget Plan'!B${i}`,
      budget: `='Budget Plan'!H${i}`,
      committed: 0, // Will be filled in as contracts are awarded
      invoiced: 0, // Will be filled in as invoices are received
      paid: 0, // Will be filled in as payments are made
      remaining: `=D${i}-E${i}`, // Committed minus Invoiced
      variance: `=C${i}-D${i}`, // Budget minus Committed
      complete: `=IF(C${i}=0,0,E${i}/C${i})`, // Percent complete based on invoices vs budget
      notes: '' // For tracking notes
    });
  }

  // Add the cost tracking data
  costTrackingSheet.addRows(costTrackingData);

  // Add a total row
  const costTrackingTotalRow = costTrackingSheet.addRow(['TOTAL', '', '=SUM(C2:C28)', '=SUM(D2:D28)', '=SUM(E2:E28)', '=SUM(F2:F28)', '=SUM(G2:G28)', '=SUM(H2:H28)', '=IF(C29=0,0,E29/C29)', '']);
  costTrackingTotalRow.font = { bold: true };

  // Apply number formatting to cost tracking sheet
  for (let i = 2; i <= costTrackingSheet.rowCount; i++) {
    for (let j = 3; j <= 8; j++) {
      costTrackingSheet.getRow(i).getCell(j).numFmt = '$#,##0.00';
    }
    if (i < costTrackingSheet.rowCount) {
      costTrackingSheet.getRow(i).getCell(9).numFmt = '0.00%';
    }
  }

  // Format with alternating row colors
  for (let i = 2; i <= costTrackingSheet.rowCount; i++) {
    if (i % 2 === 0) {
      costTrackingSheet.getRow(i).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' }
      };
    }
  }
  
  // Add borders
  for (let i = 1; i <= costTrackingSheet.rowCount; i++) {
    for (let j = 1; j <= 10; j++) {
      costTrackingSheet.getRow(i).getCell(j).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
  }

  // Create Dashboard Sheet
  const dashboardSheet = workbook.addWorksheet('Dashboard', {
    properties: { tabColor: { argb: '4BACC6' } }
  });

  // Set up columns for the dashboard
  dashboardSheet.columns = [
    { header: '', key: 'label', width: 30 },
    { header: '', key: 'value', width: 15 },
    { header: '', key: 'graph', width: 5 },
    { header: '', key: 'spacer', width: 5 },
    { header: '', key: 'percent', width: 15 }
  ];

  // Add project summary section
  dashboardSheet.mergeCells('A1:E1');
  dashboardSheet.getCell('A1').value = 'PROJECT SUMMARY';
  dashboardSheet.getCell('A1').font = { bold: true, size: 14 };
  dashboardSheet.getCell('A1').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '4BACC6' }
  };
  dashboardSheet.getCell('A1').font = { color: { argb: 'FFFFFF' }, bold: true };
  dashboardSheet.getCell('A1').alignment = { horizontal: 'center' };

  // Add budget summary
  dashboardSheet.getCell('A3').value = 'Total Budget';
  dashboardSheet.getCell('B3').value = "='Budget Plan'!H29";
  dashboardSheet.getCell('B3').numFmt = '$#,##0.00';
  
  dashboardSheet.getCell('A4').value = 'Total Committed Cost';
  dashboardSheet.getCell('B4').value = "='Cost Tracking'!D29";
  dashboardSheet.getCell('B4').numFmt = '$#,##0.00';
  
  dashboardSheet.getCell('A5').value = 'Total Invoiced to Date';
  dashboardSheet.getCell('B5').value = "='Cost Tracking'!E29";
  dashboardSheet.getCell('B5').numFmt = '$#,##0.00';
  
  dashboardSheet.getCell('A6').value = 'Total Paid to Date';
  dashboardSheet.getCell('B6').value = "='Cost Tracking'!F29";
  dashboardSheet.getCell('B6').numFmt = '$#,##0.00';
  
  dashboardSheet.getCell('A7').value = 'Budget Remaining';
  dashboardSheet.getCell('B7').value = "=B3-B4";
  dashboardSheet.getCell('B7').numFmt = '$#,##0.00';
  
  dashboardSheet.getCell('A8').value = 'Project Percent Complete';
  dashboardSheet.getCell('B8').value = "='Cost Tracking'!I29";
  dashboardSheet.getCell('B8').numFmt = '0.00%';

  // Add percentage indicators
  dashboardSheet.getCell('E4').value = '=B4/B3';
  dashboardSheet.getCell('E4').numFmt = '0.00%';
  dashboardSheet.getCell('E5').value = '=B5/B3';
  dashboardSheet.getCell('E5').numFmt = '0.00%';
  dashboardSheet.getCell('E6').value = '=B6/B3';
  dashboardSheet.getCell('E6').numFmt = '0.00%';
  dashboardSheet.getCell('E7').value = '=B7/B3';
  dashboardSheet.getCell('E7').numFmt = '0.00%';

  // Add schedule summary
  dashboardSheet.mergeCells('A10:E10');
  dashboardSheet.getCell('A10').value = 'SCHEDULE SUMMARY';
  dashboardSheet.getCell('A10').font = { bold: true, size: 14 };
  dashboardSheet.getCell('A10').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '4BACC6' }
  };
  dashboardSheet.getCell('A10').font = { color: { argb: 'FFFFFF' }, bold: true };
  dashboardSheet.getCell('A10').alignment = { horizontal: 'center' };

  // Add month headers
  const months = ['Month 1', 'Month 2', 'Month 3', 'Month 4', 'Month 5', 'Month 6', 'Month 7', 'Month 8'];
  for (let i = 0; i < months.length; i++) {
    dashboardSheet.getCell(`A${12 + i}`).value = months[i];
    dashboardSheet.getCell(`B${12 + i}`).value = `='Budget Forecast'!${String.fromCharCode(68 + i)}29`;
    dashboardSheet.getCell(`B${12 + i}`).numFmt = '$#,##0.00';
    dashboardSheet.getCell(`E${12 + i}`).value = `='Budget Forecast'!${String.fromCharCode(68 + i)}31`;
    dashboardSheet.getCell(`E${12 + i}`).numFmt = '0.00%';
  }

  // Add top categories section
  dashboardSheet.mergeCells('A21:E21');
  dashboardSheet.getCell('A21').value = 'TOP EXPENSE CATEGORIES';
  dashboardSheet.getCell('A21').font = { bold: true, size: 14 };
  dashboardSheet.getCell('A21').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '4BACC6' }
  };
  dashboardSheet.getCell('A21').font = { color: { argb: 'FFFFFF' }, bold: true };
  dashboardSheet.getCell('A21').alignment = { horizontal: 'center' };

  // Add instructions sheet
  const instructionsSheet = workbook.addWorksheet('Instructions', {
    properties: { tabColor: { argb: '808080' } }
  });
  
  // Set wider column for instructions
  instructionsSheet.columns = [
    { header: 'Construction Budget Workbook Instructions', width: 100 }
  ];
  
  // Style the header
  instructionsSheet.getRow(1).font = { bold: true, size: 14 };
  instructionsSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' }
  };
  instructionsSheet.getRow(1).font = { color: { argb: 'FFFFFF' }, bold: true };
  
  // Add instructions text
  const instructions = [
    'This workbook contains the following sheets to help manage your construction budget:',
    '',
    '1. Budget Plan: The main budget sheet where you enter all line items for your construction project.',
    '   - Enter your own categories, descriptions, quantities and unit prices',
    '   - The sheet will automatically calculate the total budget for each line item',
    '   - A contingency row is included and set to 10% of the total budget by default',
    '',
    '2. Budget Forecast: Shows the anticipated spending by month for each budget category.',
    '   - The sheet pulls data from the Budget Plan sheet',
    '   - Distribute expected costs across months by adjusting the percentage allocation for each line item',
    '   - Monthly and cumulative totals are calculated automatically',
    '   - Use this to plan cash flow needs throughout the project',
    '',
    '3. Cost Tracking: Track actual costs against budgeted amounts.',
    '   - Enter committed costs as contracts are awarded and purchase orders are issued',
    '   - Track invoiced and paid amounts',
    '   - Monitor budget variance (budgeted vs. committed)',
    '   - Track percent complete for each category',
    '   - Add notes to document changes or issues',
    '',
    '4. Dashboard: Provides a high-level overview of project budget status.',
    '   - Shows key budget metrics, including total budget, committed costs, and payments',
    '   - Displays monthly forecast and percentage complete',
    '   - Highlights top expense categories',
    '',
    'Tips for using this workbook:',
    '- Start by customizing the Budget Plan sheet with your specific line items',
    '- Update the Budget Forecast to distribute costs across the project timeline',
    '- Regularly update the Cost Tracking sheet with actual costs as they are committed and paid',
    '- Review the Dashboard for a quick project status overview',
    '- The sheet is linked with formulas, so changes in one area will update related sections automatically',
    '',
    'For best results:',
    '- Update the workbook at least weekly',
    '- Review with your project team regularly',
    '- Document changes in the Notes sections',
    '- Archive a copy of the workbook monthly to track changes over time'
  ];
  
  // Add each line of instructions as a new row
  for (let i = 0; i < instructions.length; i++) {
    instructionsSheet.addRow([instructions[i]]);
  }
  
  // Adjust row heights for better readability
  for (let i = 1; i <= instructionsSheet.rowCount; i++) {
    instructionsSheet.getRow(i).height = 18;
  }

  // Set sheet order and active sheet
  workbook.removeWorksheet('Sheet1');
  workbook.views = [
    {
      firstSheet: 0,
      activeTab: 0,
      visibility: 'visible'
    }
  ];

  // Save the workbook
  await workbook.xlsx.writeFile('Construction_Budget_Workbook.xlsx');
  
  console.log('Construction Budget Workbook created successfully!');
}

createConstructionBudgetWorkbook();