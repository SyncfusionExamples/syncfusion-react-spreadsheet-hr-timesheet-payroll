import { getCell, getColumn, getFormatFromType, getSheet, setCell, SheetDirective, SheetsDirective, Spreadsheet, SpreadsheetComponent, type CellModel, type CellStyleModel, type ChartModel, type RowModel, type SheetModel, type SortDescriptor } from '@syncfusion/ej2-react-spreadsheet';
import { attendanceRateImage, averageNetSalaryImage, averageWorkImage, dollarImage, employeeData, highLeaveEmployeeImage, highOtEmployeeImage, LateLoginImage, lowHrsEmployeeImage, OtpercentImage, overtimeHrsImage, overtimePayImage, socialContributionImage, totalDeductionImage, totalEmployeeImage, updatedTimeSheet, workingHrsImage } from './Data'

export default function App() {
  
  //initializing spreadsheet
  let spreadsheet: Spreadsheet;
  //Global variable to load data based on sheet
  let dataLoaded: string = 'EmployeeMaster';
  let dashboardSheetName: string = 'Dashboard';
  let employeeSheetName: string = 'Employee Master';
  let timeSheetName: string = 'Timesheet';
  let payrollSheetName: string = 'Payroll';
  let dataSheetName: string = 'Data';
  //Header cell style
  const headerStyle: CellStyleModel = { fontSize: '14pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#001A4A', color: '#fff', fontWeight: 'bold' };
  //created function
  const onCreated = (): void => {
    setFormats();
    setRowColumnSize();
    spreadsheet.updateRange({ dataSource: employeeData as any, startCell: 'A5' }, 1);
  }

  //datasource changed event
  const dataSourceChanged = (args: any): void => {
    if (args.action === 'dataSourceChanged') {
      if (dataLoaded === 'EmployeeMaster') {
        initiateEmployeeSheet();
        dataLoaded = 'TimeSheet';
        spreadsheet.updateRange({ dataSource: updatedTimeSheet as any, startCell: 'A5' }, 2);
      }
      else if (dataLoaded === 'TimeSheet') {
        initiateTimeSheet();
        initiatePayrollSheet();
        initiateDataSheet();
        initiateDashboardSheet();
      }
    }
  }
 
  const setRowColumnSize: Function = (): void => {
    //set the row height
    spreadsheet.setRowsHeight(30, [`${employeeSheetName}!1:57`, `${timeSheetName}!1:1280`, `${payrollSheetName}!1:56`, `${dashboardSheetName}!1:20`,]);
    spreadsheet.setRowsHeight(40, [`${dashboardSheetName}!16`, `${dashboardSheetName}!21:22`]);
    spreadsheet.setRowsHeight(35, [`${dashboardSheetName}!23:27`])
    //set the column width
    spreadsheet.setColumnsWidth(100, [`${employeeSheetName}!A:H`, `${timeSheetName}!A:K`, `${payrollSheetName}!A:M`]);
    spreadsheet.setColumnsWidth(30, [`${dashboardSheetName}!A`]);
    spreadsheet.setColumnsWidth(100, [`${dashboardSheetName}!B:Z`]);
    spreadsheet.setColumnsWidth(130, [`${employeeSheetName}!A:H`, `${timeSheetName}!A:K`, `${payrollSheetName}!A:M`]);
    //set row height
    spreadsheet.setRowsHeight(40, [`${employeeSheetName}!4:5`, `${timeSheetName}!4:5`, `${payrollSheetName}!4:5`]);
  }

  const setFormats: Function = (): void => {
    //setting sheetmodels
    const employeeSheet: SheetModel = getSheet(spreadsheet, 1);
    const timeSheet: SheetModel = getSheet(spreadsheet, 2);
    const payrollSheet: SheetModel = getSheet(spreadsheet, 3);
    //Employee Sheet Formats
    //set header value
    setCell(0, 0, employeeSheet, { value: 'EMPLOYEE MASTER SHEET- HR PAYROLL AND TIMESHEET APPLICATION', colSpan: 8, rowSpan: 3, style: headerStyle });
    //updating styles to headers
    for (let columnHeader = 0; columnHeader < 8; columnHeader++) {
      //setCell(4, columnHeader, employeeSheet, { style: columnHeaderStyle });
      if (columnHeader < 2) {
        setCell(4, columnHeader, employeeSheet, { style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#1E3A8A', color: '#fff', fontWeight: 'bold' } });
      } else if (columnHeader < 4) {
        setCell(4, columnHeader, employeeSheet, { style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#14532D', color: '#fff', fontWeight: 'bold' } });
      } else if (columnHeader < 5) {
        setCell(4, columnHeader, employeeSheet, { style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: ' #164E63', color: '#fff', fontWeight: 'bold' } });
      } else if (columnHeader < 7) {
        setCell(4, columnHeader, employeeSheet, { style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#854D0E', color: '#fff', fontWeight: 'bold' } });
      } else {
        setCell(4, columnHeader, employeeSheet, { style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#9A3412', color: '#fff', fontWeight: 'bold' } });
      }
    }
    //Time Sheet Formats
    //updating styles to timesheet headers
    for (let columnHeader = 0; columnHeader <= 10; columnHeader++) {
      //setCell(4, columnHeader, timeSheet, { style: columnHeaderStyle });
      if (columnHeader < 2) {
        setCell(4, columnHeader, timeSheet, { style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#1E3A8A', color: '#fff', fontWeight: 'bold' } });
      } else if (columnHeader < 5) {
        setCell(4, columnHeader, timeSheet, { style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#4C1D95', color: '#fff', fontWeight: 'bold' } });
      } else if (columnHeader < 8) {
        setCell(4, columnHeader, timeSheet, { style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: ' #075985', color: '#fff', fontWeight: 'bold' } });
      } else if (columnHeader < 10) {
        setCell(4, columnHeader, timeSheet, { style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#854D0E', color: '#fff', fontWeight: 'bold' } });
      } else {
        setCell(4, columnHeader, timeSheet, { style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#9A3412', color: '#fff', fontWeight: 'bold' } });
      }
    }
    //set header value
    setCell(0, 0, timeSheet, { value: 'TIME SHEET - HR PAYROLL AND TIMESHEET APPLICATION', colSpan: 11, rowSpan: 3, style: headerStyle });
    //set header value
    setCell(0, 0, payrollSheet, { value: 'PAYROLL SHEET - HR PAYROLL AND TIMESHEET APPLICATION', colSpan: 13, rowSpan: 3, style: headerStyle });
  }

  //EmployeeSheet Calculations
  const initiateEmployeeSheet: Function = (): void => {
    //get the sheet
    const employeeSheet: SheetModel = getSheet(spreadsheet, 1);
    const employeeMasterGroupHeaders: string[] = ["EMPLOYEE INFORMATION", "DESIGNATION", "WORK LOCATION", "COMPENSATION", "SHIFT HRS"];
    setCell(3, 0, employeeSheet, { value: employeeMasterGroupHeaders[0], colSpan: 2 , style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#1E40AF', color: '#fff', fontWeight: 'bold' }});
    setCell(3, 2, employeeSheet, { value: employeeMasterGroupHeaders[1], colSpan: 2 , style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#15803D', color: '#fff', fontWeight: 'bold' }});
    setCell(3, 4, employeeSheet, { value: employeeMasterGroupHeaders[2], style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#0E7490', color: '#fff', fontWeight: 'bold' } });
    setCell(3, 5, employeeSheet, { value: employeeMasterGroupHeaders[3], colSpan: 2 , style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#CA8A04', color: '#fff', fontWeight: 'bold' }});
    setCell(3, 7, employeeSheet, { value: employeeMasterGroupHeaders[4], style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#EA580C', color: '#fff', fontWeight: 'bold' } });
    //Employee ID Columns
    spreadsheet.cellFormat({ color: '#2549BE', fontWeight: 'bold' }, `${employeeSheetName}!A6:A${employeeSheet.rows.length}`);
    //department color conditional formatttings
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'HR', range: `${employeeSheetName}!C6:C${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D5DEE9' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Finance', range: `${employeeSheetName}!C6:C${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D7E7C7' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'IT', range: `${employeeSheetName}!C6:C${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D5C2E6' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Operations', range: `${employeeSheetName}!C6:C${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#DADADA' } } });
    //work location conditional formattings
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Dallas', range: `${employeeSheetName}!E6:E${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D6DDE6' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'San Francisco', range: `${employeeSheetName}!E6:E${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#F2E3B6' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Austin', range: `${employeeSheetName}!E6:E${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D7E6DB' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Seattle', range: `${employeeSheetName}!E6:E${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D9D6E6' } } });
    //set data bars to base salary
    spreadsheet.conditionalFormat({ type: 'LightBlueDataBar', range: `${employeeSheetName}!F6:F${employeeSheet.rows.length}` });
    spreadsheet.conditionalFormat({ type: 'GWRColorScale', range: `${employeeSheetName}!G6:G${employeeSheet.rows.length}` });
    //add data validation
    spreadsheet.addDataValidation({ type: 'List', value1: 'HR,IT,Finance,Operations', ignoreBlank: false }, `${employeeSheetName}!C6:C${employeeSheet.rows.length}`);
    spreadsheet.addDataValidation({ type: 'List', value1: 'Analyst,Engineer,Executive,Manager', ignoreBlank: false }, `${employeeSheetName}!D6:D${employeeSheet.rows.length}`);
    spreadsheet.addDataValidation({ type: 'List', value1: 'Dallas,San Francisco,Austin,Seattle', ignoreBlank: false }, `${employeeSheetName}!E6:E${employeeSheet.rows.length}`);
    //color formatting for cells
    spreadsheet.cellFormat({backgroundColor:'#fff'},`${employeeSheetName}!A6:H${employeeSheet.rows.length - 1}`);
    //header border
    spreadsheet.setBorder({ border: '1px solid #fff' }, `${employeeSheetName}!A1:H5`);
    spreadsheet.setBorder({ border: '1px solid #e6e6e6' }, `${employeeSheetName}!A6:H${employeeSheet.rows.length - 1}`);
    //formula update for OT rate
    for (let row = 5; row < employeeSheet.rows.length - 1; row++) {
      setCell(row, 6, employeeSheet, { formula: `=(F${row + 1}/(30*8))*1.5` });
    }
  }

  //TimeSheet Calculations
  const initiateTimeSheet: Function = (): void => {
    //get the sheet
    const timeSheet: SheetModel = getSheet(spreadsheet, 2);
    const timesheetGroupHeaders: string[] = ["ATTENDANCE IDENITIFICATION", "TIME TRACKING", "PERMISSIONS", "WORK SUMMARY", "LEAVE"];
    setCell(3, 0, timeSheet, { value: timesheetGroupHeaders[0], colSpan:2, style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#1E40AF', color: '#fff', fontWeight: 'bold' }});
    setCell(3, 2, timeSheet, { value: timesheetGroupHeaders[1], colSpan:3, style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#6D28D9', color: '#fff', fontWeight: 'bold' }});
    setCell(3, 5, timeSheet, { value: timesheetGroupHeaders[2], colSpan:3, style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#0EA5E9', color: '#fff', fontWeight: 'bold' }});
    setCell(3, 8, timeSheet, { value: timesheetGroupHeaders[3], colSpan:2, style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#CA8A04', color: '#fff', fontWeight: 'bold' }});
    setCell(3, 10, timeSheet, { value: timesheetGroupHeaders[4], style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#EA580C', color: '#fff', fontWeight: 'bold' }});
    //leave
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'CL', range: `${timeSheetName}!K6:K${timeSheet.rows.length}`, format: { style: { backgroundColor: '#4CAF50' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'LOP', range: `${timeSheetName}!K6:K${timeSheet.rows.length}`, format: { style: { backgroundColor: '#ff0000' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'SL', range: `${timeSheetName}!K6:K${timeSheet.rows.length}`, format: { style: { backgroundColor: '#ff7003' } } });
    //conditional formattings
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '9:30', range: `${timeSheetName}!C6:C${timeSheet.rows.length}`, format: { style: { color: '#ff0000', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'LessThan', value: '9:31', range: `${timeSheetName}!C6:C${timeSheet.rows.length}`, format: { style: { color: '#4CAF50', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'LessThan', value: '17:31', range: `${timeSheetName}!D6:D${timeSheet.rows.length}`, format: { style: { color: '#ff0000', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '17:30', range: `${timeSheetName}!D6:D${timeSheet.rows.length}`, format: { style: { color: '#4CAF50', fontWeight: 'bold' } } });
    //Lunch
    spreadsheet.cellFormat({backgroundColor:'#f1deab'},`${timeSheetName}!E6:E${timeSheet.rows.length}`);
    //late hrs
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '0', range: `${timeSheetName}!F6:F${timeSheet.rows.length}`, format: { style: { color: '#FF0000', fontWeight: 'bold' } } });
    //Permission
    spreadsheet.conditionalFormat({ type: 'LessThan', value: '1:01', range: `${timeSheetName}!G6:G${timeSheet.rows.length}`, format: { style: { color: '#4CAF50', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'Between', value: '1:01,2:01', range: `${timeSheetName}!G6:G${timeSheet.rows.length}`, format: { style: { color: '#FFC107', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '2:00', range: `${timeSheetName}!G6:G${timeSheet.rows.length}`, format: { style: { color: '#FF0000', fontWeight: 'bold' } } });
    //Break
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '0.014', range: `${timeSheetName}!H6:H${timeSheet.rows.length}`, format: { style: { color: '#FF0000', fontWeight: 'bold' } }});
    //work hours
    spreadsheet.conditionalFormat({ type: 'LightBlueDataBar', range: `${timeSheetName}!I6:I${timeSheet.rows.length}` });
    //overtime hours
    spreadsheet.conditionalFormat({ type: 'OrangeDataBar', range: `${timeSheetName}!J6:J${timeSheet.rows.length}` });
    //formula update for permissions, work hours and overtime
    for (let row = 5; row <= timeSheet.rows.length - 1; row++) {
      //permission formula
      setCell(row, 5, timeSheet, { formula: `=IF(C${row + 1}="-","-",IF(C${row + 1}>TIME(9,30,0),C${row + 1} - TIME(9,30,0),0))`, format: 'h:mm' });
      //workhour formula
      setCell(row, 8, timeSheet, { formula: `=IF(C${row + 1}="-","-",(D${row + 1}-C${row + 1})-(E${row + 1}+F${row + 1}+G${row + 1}+H${row + 1}))`, format: 'h:mm' });
      //overtime formula
      setCell(row, 9, timeSheet, { formula: `=IF(I${row + 1}="-","-",IF(I${row + 1}>TIME(7,30,0),I${row + 1}-TIME(7,30,0),0))`, format: 'h:mm' });
      //LOP formula update
      if (timeSheet.rows[row].cells[10].value == "None") {
        setCell(row, 10, timeSheet, { formula: `=IF(G${row + 1}>0,IF(SUMIFS(G6:G${row},B6:B${row},B${row + 1})>=0.125,"LOP","None"),"None")` });
      }
    }
    //add data validation
    spreadsheet.addDataValidation({ type: 'List', value1: 'None,SL,LOP,CL', ignoreBlank: false }, `${timeSheetName}!K6:K${timeSheet.rows.length}`);
    //applying border values
    spreadsheet.setBorder({ border: '1px solid #f2f2f2' }, `${timeSheetName}!A6:J${timeSheet.rows.length}`, 'Horizontal');
    spreadsheet.setBorder({ border: '1px solid #f2f2f2' }, `${timeSheetName}!A6:J${timeSheet.rows.length}`, 'Vertical');
    spreadsheet.setBorder({ border: '1px solid #fff' }, `${timeSheetName}!A1:J5`);
    //color formatting for cells
    spreadsheet.numberFormat('d/m/yyyy',`${timeSheetName}!A6:A${timeSheet.rows.length}`);
    spreadsheet.cellFormat({backgroundColor:'#fff'},`${timeSheetName}!A6:D${timeSheet.rows.length} F6:K${timeSheet.rows.length}`);
  }

  //Payroll Calculations
  const initiatePayrollSheet: Function = (): void => {
    //get the employee sheet
    const employeeSheet: SheetModel = getSheet(spreadsheet, 1);
    //get the time sheet
    const timeSheet: SheetModel = getSheet(spreadsheet, 2);
    //get the payroll sheet
    const payrollSheet: SheetModel = getSheet(spreadsheet, 3);
    const payrollGroupHeaders: string[] = ["EMPLOYEE PAY INFO", "WORK HRS AND PAY", "LEAVE AND DEDUCTIONS", "STATUATORY CONTRIBUTIONS", "TAX DEDUCTIONS", "NET PAY"];
    setCell(3, 0, payrollSheet, { value: payrollGroupHeaders[0], colSpan: 3, style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#1E40AF', color: '#fff', fontWeight: 'bold' }  });
    setCell(3, 3, payrollSheet, { value: payrollGroupHeaders[1], colSpan: 3, style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#16A34A', color: '#fff', fontWeight: 'bold' } });
    setCell(3, 6, payrollSheet, { value: payrollGroupHeaders[2], colSpan: 2, style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#EA580C', color: '#fff', fontWeight: 'bold' } });
    setCell(3, 8, payrollSheet, { value: payrollGroupHeaders[3], colSpan: 2, style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#4F46E5', color: '#fff', fontWeight: 'bold' } });
    setCell(3, 10, payrollSheet, { value: payrollGroupHeaders[4], colSpan: 2, style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#DC2626', color: '#fff', fontWeight: 'bold' } });
    setCell(3, 12, payrollSheet, { value: payrollGroupHeaders[5], style:{ fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#22C55E', color: '#fff', fontWeight: 'bold' } });
    //payroll sheet header update
    const payrollHeaders: string[] = ['Employee Id', 'Department', 'Base Salary', 'Total Hours', 'OT Hours', 'OT Pay', 'Leave Days', 'Leave Deduction', 'Social\nContribution %', 'Social\nContribution', 'Tax%', 'Tax', 'Net Salary'];
    payrollHeaders.forEach((value, index) => {
      if (index < 3) {
        setCell(4, index, payrollSheet, { value: value, style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#1E3A8A', color: '#fff', fontWeight: 'bold' } });
      } else if(index < 6){
        setCell(4, index, payrollSheet, { value: value, style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#166534', color: '#fff', fontWeight: 'bold' } });
      } else if (index < 8){
        setCell(4, index, payrollSheet, { value: value, style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: ' #9A3412', color: '#fff', fontWeight: 'bold' } });
      } else if (index < 10){
        setCell(4, index, payrollSheet, { value: value, style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#312E81', color: '#fff', fontWeight: 'bold' } });      
      } else if (index < 12){
        setCell(4, index, payrollSheet, { value: value, style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#7F1D1D', color: '#fff', fontWeight: 'bold' } });      
      } else{
        setCell(4, index, payrollSheet, { value: value, style: { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#15803D', color: '#fff', fontWeight: 'bold' } });
      }
      });
    //formula update
    for (let row = 5; row < employeeSheet.rows.length - 1; row++) {
      //employee id
      setCell(row, 0, payrollSheet, { formula: `=${employeeSheetName}!A${row + 1}`, style: { color: '#2549BE', fontWeight: 'bold' } });
      //department
      setCell(row, 1, payrollSheet, { formula: `=${employeeSheetName}!C${row + 1}`, style: { fontWeight: 'bold' } });
      //base salary
      setCell(row, 2, payrollSheet, { formula: `=VLOOKUP(A${row + 1},${employeeSheetName}!A1:G56,6,FALSE)`, format: getFormatFromType('Currency') });
      //total hours
      setCell(row, 3, payrollSheet, { formula: `=SUMIFS(${timeSheetName}!I1:I${timeSheet.rows.length},${timeSheetName}!B1:B${timeSheet.rows.length},A${row + 1})`, format: '[h]:mm' });
      //ot hours
      setCell(row, 4, payrollSheet, { formula: `=SUMIFS(Timesheet!J6:J${timeSheet.rows.length},Timesheet!B6:B${timeSheet.rows.length},A${row + 1})`, format: '[h]:mm' });
      //ot pay
      setCell(row, 5, payrollSheet, { formula: `=E${row + 1}*24*((C${row + 1}/240)*1.5)`, style: { color: '#1F3A6D', fontWeight: 'bold' } });
      //leave days
      setCell(row, 6, payrollSheet, { formula: `=COUNTIFS(Timesheet!K6:K${timeSheet.rows.length},"LOP",Timesheet!B6:B${timeSheet.rows.length},A${row + 1})` });
      //leave deduction
      setCell(row, 7, payrollSheet, { formula: `=(C${row + 1}/30)*G${row + 1}` });
      //social contribution %
      if (row === 7 || row === 9 || row === 11 || row === 14 || row === 17 || row === 19 || row === 25 || row === 37 || row === 42) {
        setCell(row, 8, payrollSheet, { value: '0.13', format: `${getFormatFromType('Percentage')}`, style: { color: '#2F6FCC', fontWeight: 'bold' } });
      } else {
        setCell(row, 8, payrollSheet, { value: '0.12', format: `${getFormatFromType('Percentage')}`, style: { color: '#2F6FCC', fontWeight: 'bold' } });
      }
      //Contribution
      setCell(row, 9, payrollSheet, { formula: `=C${row + 1}*I${row + 1}`, format: `${getFormatFromType('Currency')}`, style: { color: '#2F6FCC', fontWeight: 'bold' } });
      //Tax %
      setCell(row, 10, payrollSheet, { formula: `=IF(C${row + 1}*12<=17700,0.10,IF(C${row + 1}*12<=67450,0.12,IF(C${row + 1}*12<=105700,0.22,IF(C${row + 1}*12<=201750,0.24,IF(C${row + 1}*12<=256200,0.32,IF(C${row + 1}*12<=640600,0.35,0.37))))))`, format: `${getFormatFromType('Percentage')}`, style: { backgroundColor: '#EFE6D8', color: '#D35400', fontWeight: 'bold' } });
      //Tax
      setCell(row, 11, payrollSheet, { formula: `=C${row + 1}*K${row + 1}`, format: `${getFormatFromType('Currency')}`, style: { backgroundColor: '#EFE6D8', color: '#D35400', fontWeight: 'bold' } });
      //net salary
      setCell(row, 12, payrollSheet, { formula: `=C${row + 1}+F${row + 1}-(H${row + 1}+J${row + 1}+L${row + 1})`, style: { fontWeight: 'bold' } });
    }
    //Data bar
    spreadsheet.conditionalFormat({ type: 'GreenDataBar', range: `${payrollSheetName}!C6:C${payrollSheet.rows.length}` });
    //Leave Days
    spreadsheet.conditionalFormat({ type: 'RYGColorScale', range: `${payrollSheetName}!G6:G${payrollSheet.rows.length}` });
    //net salary
    spreadsheet.conditionalFormat({ type: 'GWColorScale', range: `${payrollSheetName}!M6:M${payrollSheet.rows.length}` });
    //Leave Deduction
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '0', range: `${payrollSheetName}!H6:H${payrollSheet.rows.length}`, format: { style: { color: '#FF0000', fontWeight: 'bold' } }});
    //PF
    spreadsheet.setBorder({ border: '1px solid #e6e6e6' }, `${payrollSheetName}!A6:M${payrollSheet.rows.length}`);
    //header border
    spreadsheet.setBorder({ border: '1px solid #fff' }, `${payrollSheetName}!A1:M5`);
    //department color conditional formatttings
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'HR', range: `${payrollSheetName}!B6:B${payrollSheet.rows.length}`, format: { style: { backgroundColor: '#D5DEE9' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Finance', range: `${payrollSheetName}!B6:B${payrollSheet.rows.length}`, format: { style: { backgroundColor: '#D7E7C7' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'IT', range: `${payrollSheetName}!B6:B${payrollSheet.rows.length}`, format: { style: { backgroundColor: '#D5C2E6' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Operations', range: `${payrollSheetName}!B6:B$${payrollSheet.rows.length}`, format: { style: { backgroundColor: '#DADADA' } } });
    //add data validation
    spreadsheet.addDataValidation({ type: 'List', value1: 'HR,IT,Finance,Operations', ignoreBlank: false }, `${payrollSheetName}!B6:B${payrollSheet.rows.length}`);
    spreadsheet.cellFormat({backgroundColor:'#fff'},`${payrollSheetName}!A6:J${payrollSheet.rows.length}`);
  }

  //DataSheet Calculations
  const initiateDataSheet: Function = (): void => {
    //get employee sheet model
    const employeeSheet: SheetModel = getSheet(spreadsheet, 1);
    const employeeSheetLastRow: number = employeeSheet.rows.length;
    //get timesheet model
    const timeSheet: SheetModel = getSheet(spreadsheet, 2);
    const timeSheetLastRow: number = timeSheet.rows.length;
    //get payroll sheet model
    const payrollSheet: SheetModel = getSheet(spreadsheet, 3);
    const payrollSheetLastRow: number = payrollSheet.rows.length;
    //get data sheet model
    const dataSheet: SheetModel = getSheet(spreadsheet, 4);
    //dept & payroll cost
    setCell(2, 0, dataSheet, { value: 'Department' });
    setCell(2, 1, dataSheet, { value: 'Payroll Cost' });
    setCell(8, 0, dataSheet, { value: 'Department' });
    setCell(8, 1, dataSheet, { value: 'OT Cost' });
    setCell(8, 2, dataSheet, { value: 'Avg Work Hrs' });
    setCell(8, 3, dataSheet, { value: 'Employees' });
    setCell(18, 0, dataSheet, { value: 'Department' });
    setCell(18, 1, dataSheet, { value: 'Average Risk Score' });
    //populating data for payroll cost and OT hours by department for charts
    const department: string[] = ['IT','HR','Finance','Operations'];
    department.forEach((value: string, index: number) => {
      setCell(3 + index, 0, dataSheet, { value: value });
      //formula for the standard payroll cost
      setCell(3 + index, 1, dataSheet, { formula:`=SUMIF(${payrollSheetName}!B6:B${payrollSheetLastRow},A${4 + index },${payrollSheetName}!M6:M${payrollSheetLastRow})` , format:getFormatFromType('Currency')});
      setCell(9 + index, 0, dataSheet, { value: value });
      setCell(19 + index, 0, dataSheet, { value: value });
    });
    //OT Hrs Distribution
    const otHrs: string[] = ['<2Hrs','2-4Hrs','>4Hrs'];
    otHrs.forEach((value: string, index: number) => {
      setCell(3 + index, 2, dataSheet, { value: value });
    });
    //othrs
    setCell(3, 3, dataSheet, {
      formula: `=COUNTIFS(Data!I6:I${employeeSheet.rows.length - 1}, "<0.0833")`
    });
    setCell(4, 3, dataSheet, {
      formula: `=COUNTIFS(Data!I6:I${employeeSheet.rows.length - 1}, ">=0.0833", Data!I6:I${employeeSheet.rows.length - 1}, "<=0.1667")`
    });
    setCell(5, 3, dataSheet, {
      formula: `=COUNTIFS(Data!I6:I${employeeSheet.rows.length - 1}, ">0.1667")`
    });

    //formulas for ot hrs
    setCell(9, 1, dataSheet, { formula: `=SUMIF(${payrollSheetName}!B6:B${payrollSheetLastRow},A10, ${payrollSheetName}!F6:F${payrollSheetLastRow})`, format:'$#,##0_);($#,##0)'});
    setCell(10, 1, dataSheet, { formula: `=SUMIF(${payrollSheetName}!B6:B${payrollSheetLastRow},A11,${payrollSheetName}!F6:F${payrollSheetLastRow})`, format:'$#,##0_);($#,##0)'});
    setCell(11, 1, dataSheet, { formula: `=SUMIF(${payrollSheetName}!B6:B${payrollSheetLastRow},A12,${payrollSheetName}!F6:F${payrollSheetLastRow})`, format:'$#,##0_);($#,##0)'});
    setCell(12, 1, dataSheet, { formula: `=SUMIF(${payrollSheetName}!B6:B${payrollSheetLastRow},A13,${payrollSheetName}!F6:F${payrollSheetLastRow})`, format:'$#,##0_);($#,##0)'});
    //Employee Average working %
    const employeeAverageWorking = ['<7Hrs','7-7.5Hrs','>7.5Hrs'];
    employeeAverageWorking.forEach((value: string, index: number) => {
      setCell(9 + index, 2, dataSheet, { value: value });
    });

    //formulas for average working %
    setCell(9, 3, dataSheet, { formula: `=COUNTIFS(H6:H${employeeSheetLastRow - 1}, "<"&7/24)` });
    setCell(10, 3, dataSheet, { formula: `=COUNTIFS(H6:H${employeeSheetLastRow - 1}, ">="&7/24, H6:H${employeeSheetLastRow - 1}, "<="&7.5/24)` });
    setCell(11, 3, dataSheet, { formula: `=COUNTIFS(H6:H${employeeSheetLastRow - 1}, ">"&7.5/24)` });
    //Employee Data Summarized
    const employeeMetric = ['EmpId','Name','Department','Avg Work Hrs','Total OT Hrs','Leaves Taken','Late Login','Risk Score'];
    employeeMetric.forEach((value: string, index: number) => {
      setCell(4, 4 + index, dataSheet, { value: value });
    });

    for(let employeeRow = 5 ; employeeRow < employeeSheetLastRow - 1 ; employeeRow++){
      //get employeename
       setCell(employeeRow, 4, dataSheet, { formula:`=${employeeSheetName}!B${employeeRow + 1}` });
      //get employeeID
      setCell(employeeRow, 5, dataSheet, { formula:`=${employeeSheetName}!A${employeeRow + 1}` });
      //get department
      setCell(employeeRow, 6, dataSheet, { formula:`=${employeeSheetName}!C${employeeRow + 1}` });
      //get avg work hrs
      setCell(employeeRow, 7, dataSheet, { formula: `=IFERROR( VLOOKUP(F${employeeRow + 1}, ${payrollSheetName}!A6:D${payrollSheetLastRow}, 4, FALSE) / COUNTIFS(Timesheet!B6:B${timeSheetLastRow}, F${employeeRow + 1}), 0)`, format: 'h:mm' });
      //get total ot hrs
      setCell(employeeRow, 8, dataSheet, { formula: `=VLOOKUP(F${employeeRow + 1}, ${payrollSheetName}!A6:E${payrollSheetLastRow}, 5, FALSE)`, format: 'h:mm' });
      //get leaves taken
      setCell(employeeRow, 9, dataSheet, { formula:`=COUNTIFS(${timeSheetName}!B6:B${timeSheetLastRow},F${employeeRow + 1},${timeSheetName}!K6:K${timeSheetLastRow},"<>None")` });
      //get late logins
      setCell(employeeRow, 10, dataSheet, { formula:`=COUNTIFS(${timeSheetName}!B6:B${timeSheetLastRow},F${employeeRow + 1},${timeSheetName}!F6:F${timeSheetLastRow},">0")` });
      //get risk score
      setCell(employeeRow, 11, dataSheet, { formula: `=IF(I${employeeRow + 1}*24>5,25,0) + IF(H${employeeRow + 1}*24<7.5,30,0) +IF(J${employeeRow + 1}>2,25,0)+IF(K${employeeRow + 1}>3,20,0)`, format: '0' });
    }
    //Leave Distribution
    const leaveDays: string[] = ['Casual Leave', 'Sick Leave', 'Loss of Pay'];
    leaveDays.forEach((value: string, index: number) => {
      setCell(14 + index, 0, dataSheet, { value: value });
    });
    //formulas for leave
    setCell(14, 1 , dataSheet, { formula:`=COUNTIFS(${timeSheetName}!K6:K${timeSheetLastRow},"CL")` });
    setCell(15, 1 , dataSheet, { formula:`=COUNTIFS(${timeSheetName}!K6:K${timeSheetLastRow},"SL")` });
    setCell(16, 1 , dataSheet, { formula:`=COUNTIFS(${timeSheetName}!K6:K${timeSheetLastRow},"LOP")` });
    //formulas for dept wise risk
    for (let rowNumber: number = 19; rowNumber < 23; rowNumber++) {
      setCell(rowNumber, 1, dataSheet, { formula: `=AVERAGEIFS(L6:L${dataSheet.rows.length}, G6:G${dataSheet.rows.length}, A${rowNumber + 1})`, format: '0' });
    }
  }

  const initiateDashboardSheet: Function = (): void => {
    //get the dashboard sheet
    const dashboardSheet = getSheet(spreadsheet, 0);
    //get the employee sheet
    const employeeSheet = getSheet(spreadsheet, 1);
    //get the timesheet
    const timeSheet = getSheet(spreadsheet, 2);
    //get the payroll sheet
    const payrollSheet = getSheet(spreadsheet, 3);
    //get the data sheet
    const dataSheet = getSheet(spreadsheet, 4);
    const keyMetrics: string[] = ['Total Worked Hours','Total Payroll Cost' , 'Total OT Hours','Total OT Cost','Avg Work Hrs/Day', 'Attendance Rate', 'Total Employees'];
    const keyHeaders: string[] = ['WORKFORCE EFFICIENCY','ATTENDANCE RISK','PAYROLL SUMMARY', 'COST DRIVER'];
    const employeeRiskSummary: string[] = ['Emp ID','Name','Department','Total OT(hrs)','Avg Work/day','Leave Days','Late Logins','Risk Score','Risk Level','Key Issue'];
    const keyMetricsValues: string[] = ['Low Hour Employees','High OT Employees','High Leave Employees','Frequent Late Logins','Avg Net Salary','Social Contribution','Total Deduction', 'OT Cost % of Payroll'];
    const headerStyle: CellStyleModel = { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', textIndent:'30px' };
    const subHeaderValueStyle: CellStyleModel = { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold' };
    const subHeaderStyle: CellStyleModel = { textAlign: 'center', verticalAlign: 'middle', fontSize: '10pt', backgroundColor: '#fff'};
    const headerTileStyle: CellStyleModel = { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#7A3DB5', fontWeight:'bold' };
    const employeeDataStyle: CellStyleModel = { textAlign: 'center', verticalAlign: 'middle', fontSize: '11pt', backgroundColor:'#fff', color:'#2F3E73', fontWeight:'bold' };
    //dashboard sheet cell format
    spreadsheet.cellFormat({backgroundColor:'#fff'},`${dashboardSheetName}!L19:Q19 S19:U19`);
    setCell(1, 1, dashboardSheet, { value: keyMetrics[0], colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#23A26D', fontWeight:'bold', textIndent:'30px' }});
    setCell(1, 4, dashboardSheet, { value: keyMetrics[1], colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#2E80C7', fontWeight:'bold', textIndent:'30px' }});
    setCell(1, 7, dashboardSheet, { value: keyMetrics[2], colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#EF4B4B', fontWeight:'bold', textIndent:'30px' }});
    setCell(1, 10, dashboardSheet, { value: keyMetrics[3], colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#F57C1F', fontWeight:'bold', textIndent:'30px' }});
    setCell(1, 13, dashboardSheet, { value: keyMetrics[4], colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#33A9C9', fontWeight:'bold', textIndent:'30px' }});
    setCell(1, 16, dashboardSheet, { value: keyMetrics[5], colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:' #3BB054', fontWeight:'bold', textIndent:'30px' } });
    setCell(1, 19, dashboardSheet, { value: keyMetrics[6], colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#B54AD6', fontWeight:'bold', textIndent:'30px' } });
    //populating data tiles
    let index: number = 0;
    setCell(15, 1, dashboardSheet, { value: keyHeaders[0], colSpan: 4, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '14pt', backgroundColor: '#1E5CCB', fontWeight: 'bold', color:'#fff'} });
    setCell(15, 6, dashboardSheet, { value: keyHeaders[1], colSpan: 4, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '14pt', backgroundColor: '#2E8B57', fontWeight: 'bold', color:'#fff' } });
    setCell(15, 11, dashboardSheet, { value: keyHeaders[2], colSpan: 6, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '14pt', backgroundColor: '#7A3DB5', fontWeight: 'bold', color:'#fff' } });
    setCell(15, 18, dashboardSheet, { value: keyHeaders[3], colSpan: 3, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '14pt', backgroundColor: '#F97316', fontWeight: 'bold', color:'#fff' } });
    setCell(16, 1, dashboardSheet, { value: keyMetricsValues[0], colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color: '#F39C12' } });
    setCell(16, 3, dashboardSheet, { value: keyMetricsValues[1], colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color: '#ff0000' } });
    setCell(16, 6, dashboardSheet, { value: keyMetricsValues[2], colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color: '#ff0000' } });
    setCell(16, 8, dashboardSheet, { value: keyMetricsValues[3], colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color: '#F39C12' } });
    setCell(16, 11, dashboardSheet, { value: keyMetricsValues[4], colSpan: 2, style: headerTileStyle });
    setCell(16, 13, dashboardSheet, { value: keyMetricsValues[5], colSpan: 2, style: headerTileStyle });
    setCell(16, 15, dashboardSheet, { value: keyMetricsValues[6], colSpan: 2, style: headerTileStyle });
    setCell(16, 18, dashboardSheet, { value: keyMetricsValues[7], colSpan: 3, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color: '#EA580C' }  });
    //adding formulas to the data sheet
    //total worked hrs
    setCell(2, 1, dashboardSheet, { formula: `=SUM(${payrollSheetName}!D6:D${payrollSheet.rows.length})`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#23A26D', fontWeight:'bold', textIndent:'30px' }});
    //total payroll cost
    setCell(2, 4, dashboardSheet, { formula: `=SUM(${payrollSheetName}!M6:M${payrollSheet.rows.length})`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#2E80C7', fontWeight:'bold', textIndent:'30px' } });
    //total ot hrs
    setCell(2, 7, dashboardSheet, { formula: `=SUM(${payrollSheetName}!E6:E${payrollSheet.rows.length})`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#EF4B4B', fontWeight:'bold', textIndent:'30px' } });
    //total ot pay
    setCell(2, 10, dashboardSheet, { formula: `=SUM(${payrollSheetName}!F6:F${payrollSheet.rows.length})`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#F57C1F', fontWeight:'bold', textIndent:'30px' } });
    //avg work hr/day
    setCell(2, 13, dashboardSheet, { formula: `=AVERAGE(${dataSheetName}!H6:H${dataSheet.rows.length})`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#33A9C9', fontWeight:'bold', textIndent:'30px' } });
    //attendance rate
    setCell(2, 16, dashboardSheet, { formula: `=COUNTIFS(${timeSheetName}!K6:K${timeSheet.rows.length},"None")/COUNTA(${timeSheetName}!K6:K${timeSheet.rows.length})`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor:'#fff', color:'#3BB054', fontWeight:'bold', textIndent:'30px' }, format:getFormatFromType('Percentage') });
    //totl employees
    setCell(2, 19, dashboardSheet, { formula: `=COUNTA(${employeeSheetName}!A6:A${employeeSheet.rows.length})`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', color: '#B54AD6', fontWeight: 'bold', textIndent: '30px' } });
    //low hour employees
    setCell(17, 1, dashboardSheet, { formula: `=COUNTIFS(Data!H6:H${dataSheet.rows.length},"<" &(7/24))`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color:'#F39C12' } });
    //lowhr - data
    setCell(18, 1, dashboardSheet, { value: '<7 hrs/day', colSpan: 2, style: subHeaderStyle, image: lowHrsEmployeeImage });
    //high ot employees
    setCell(17, 3, dashboardSheet, { formula: `=COUNTIFS(Data!I6:I${dataSheet.rows.length},">" &(5/24))`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color:'#ff0000' } });
    //high ot data
    setCell(18, 3, dashboardSheet, { value: '>5 hrs OT/month', colSpan: 2, style: subHeaderStyle, image: highOtEmployeeImage });
    //high leave employees
    setCell(17, 6, dashboardSheet, { formula: `=COUNTIFS(Data!J6:J${dataSheet.rows.length},">2")`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color:'#ff0000' } });
    //high leave data
    setCell(18, 6, dashboardSheet, { value: '>5 leave days', colSpan: 2, style: subHeaderStyle, image: highLeaveEmployeeImage });
    //frequent late login
    setCell(17, 8, dashboardSheet, { formula: `=COUNTIFS(Data!K6:K${dataSheet.rows.length},">5")`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color:'#F39C12' } });
    //frequent login data
    setCell(18, 8, dashboardSheet, { value: '>5 late logins', colSpan: 2, style: subHeaderStyle, image: LateLoginImage });
    //average net salary
    setCell(17, 11, dashboardSheet, { formula: `=AVERAGE(${payrollSheetName}!M6:M${payrollSheet.rows.length})`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color:'#7A3DB5' } });
    setCell(18, 11, dashboardSheet, { value: 'Avg Employee Salary', colSpan: 2, style: subHeaderStyle, image: averageNetSalaryImage });
    //pf contribution
    setCell(17, 13, dashboardSheet, { formula: `=SUM(${payrollSheetName}!J6:J${payrollSheet.rows.length})`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color:'#7A3DB5' } });
    setCell(18, 13, dashboardSheet, { value: 'Overall Social Contribution', colSpan: 2, style: subHeaderStyle, image: socialContributionImage });
    //total deduction
    setCell(17, 15, dashboardSheet, { formula: `=SUM(${payrollSheetName}!H6:H${payrollSheet.rows.length}) + SUM(${payrollSheetName}!J6:J${payrollSheet.rows.length})`, colSpan: 2, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color:'#7A3DB5' } });
    setCell(18, 15, dashboardSheet, { value: 'Social Contribution + Tax', colSpan: 2, style: subHeaderStyle, image: totalDeductionImage });
    //OT cost % of payroll
    setCell(17, 18, dashboardSheet, { formula: `=SUM(${payrollSheetName}!F6:F${payrollSheet.rows.length})/SUM(${payrollSheetName}!M6:M${payrollSheet.rows.length})`, colSpan: 3, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', fontWeight: 'bold', color:'#F57C1F' } });
    setCell(18, 18, dashboardSheet, { formula: `=IF(S18>0.10,"Above Ideal 10%","Below Ideal 10%")`, colSpan: 3, style: subHeaderStyle, image: OtpercentImage });
    setCell(20, 1, dashboardSheet, { value: 'TOP EMPLOYEE RISK SUMMARY', colSpan: 13, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '14pt', backgroundColor: '#1F2A44', color: '#fff', fontWeight: 'bold' } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'High', range: `${dashboardSheetName}!J23:J27`, format: { style: { color: '#ff0000', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Medium', range: `${dashboardSheetName}!J23:J27`, format: { style: { color: '#F57C1F', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Low', range: `${dashboardSheetName}!J23:J27`, format: { style: { color: '#4CAF50', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'RYGColorScale', range: `${dashboardSheetName}!I23:I27` });
    //employee risk data
    index = 0;
    for (let colNumber: number = 1; colNumber <= 10; colNumber++) {
      if (colNumber == 10) {
        setCell(21, colNumber, dashboardSheet, { value: employeeRiskSummary[index as number], style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', color: '#2F5BD3', fontWeight: 'bold' }, colSpan: 4 });
      } else {
        setCell(21, colNumber, dashboardSheet, { value: employeeRiskSummary[index as number], style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '12pt', backgroundColor: '#fff', color: '#2F5BD3', fontWeight: 'bold' } });
      }
      index++;
    }
    setTimeout(() => {
      //custom sort
      let sortDescriptors: SortDescriptor = {
        field: 'B',
        order: 'Descending'
      }
      spreadsheet.sort({ containsHeader: false, sortDescriptors: sortDescriptors }, 'Data!A4:B7')
        .then(() => {
          spreadsheet.sort({ containsHeader: true, sortDescriptors: { field: 'L', order: 'Descending' } }, 'Data!E5:L56').then(() => {
            //reset index
            index = 0;
            let rowNumber = 22;
            for (let colNumber: number = 0; colNumber < 5; colNumber++) {
              setCell(22 + colNumber, 1, dashboardSheet, { formula: `=Data!F${6 + colNumber}`, style: employeeDataStyle });
              setCell(22 + colNumber, 2, dashboardSheet, { formula: `=Data!E${6 + colNumber}`, style: employeeDataStyle });
              setCell(22 + colNumber, 3, dashboardSheet, { formula: `=Data!G${6 + colNumber}`, style: employeeDataStyle });
              setCell(22 + colNumber, 4, dashboardSheet, { formula: `=Data!I${6 + colNumber}`, style: employeeDataStyle });
              setCell(22 + colNumber, 5, dashboardSheet, { formula: `=Data!H${6 + colNumber}`, style: employeeDataStyle });
              setCell(22 + colNumber, 6, dashboardSheet, { formula: `=Data!J${6 + colNumber}`, style: employeeDataStyle });
              setCell(22 + colNumber, 7, dashboardSheet, { formula: `=Data!K${6 + colNumber}`, style: employeeDataStyle });
              setCell(22 + colNumber, 8, dashboardSheet, { formula: `=Data!L${6 + colNumber}`, style: employeeDataStyle });
              setCell(22 + colNumber, 9, dashboardSheet, { formula: `=IF(I${rowNumber + 1}>70,"High",IF(I${rowNumber + 1}>=50,"Medium","Low"))`, style: employeeDataStyle });
              if (colNumber < 3) {
                setCell(22 + colNumber, 10, dashboardSheet, { formula: `=CONCAT( IF(E${rowNumber + 1}>5/24,"High OT"&IF(OR(G${rowNumber + 1}>2,H${rowNumber + 1}>3,F${rowNumber + 1}<7.5/24)," + ",""),""),IF(G${rowNumber + 1}>2,"High Leave"&IF(OR(H${rowNumber + 1}>3,F${rowNumber + 1}<7/24)," + ",""),""),
                IF(H${rowNumber + 1}>3,&"Frequent Late"&IF(F${rowNumber + 1}<7.5/24," + ",""),""), IF(F${rowNumber + 1}<7.5/24,&"Low Work Hours",""))`, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '11pt', backgroundColor: '#fff', color: '#ff0000', fontWeight: 'bold' }, colSpan: 4 });
              } else {
                setCell(22 + colNumber, 10, dashboardSheet, { formula: `=CONCAT( IF(E${rowNumber + 1}>5/24,"High OT"&IF(OR(G${rowNumber + 1}>2,H${rowNumber + 1}>3,F${rowNumber + 1}<7.5/24)," + ",""),""),IF(G${rowNumber + 1}>2,"High Leave"&IF(OR(H${rowNumber + 1}>3,F${rowNumber + 1}<7/24)," + ",""),""),
                IF(H${rowNumber + 1}>3,&"Frequent Late"&IF(F${rowNumber + 1}<7.5/24," + ",""),""), IF(F${rowNumber + 1}<7.5/24,&"Low Work Hours",""))`, style: { textAlign: 'center', verticalAlign: 'middle', fontSize: '11pt', backgroundColor: '#fff', color: '#ffc000', fontWeight: 'bold' }, colSpan: 4 });
              }
              index++;
              rowNumber++;
            }
            spreadsheet.numberFormat('0.00%',`${dashboardSheetName}!S18:U18`);
            //setting border
            spreadsheet.setBorder({ border: '1px solid #e6e6e6' }, `${dashboardSheetName}!B22:N27`);
            spreadsheet.setBorder({ border: '1px solid #666666' }, `${dashboardSheetName}!B16:E16 G16:J16 L16:Q16`,'Outer');
            spreadsheet.setBorder({ border: '1px solid #e6e6e6' }, `${dashboardSheetName}!B2:C3 E2:F3 H2:I3 K2:L3 N2:O3 Q2:R3 T2:U3 B17:E19 G17:J19 L17:Q19 S17:U19`, 'Outer');
            //applying sheet protection
            spreadsheet.protectSheet(1, { formatCells: false, formatRows: false, formatColumns: false, insertLink: false, selectCells: true });
            spreadsheet.protectSheet(2, { formatCells: false, formatRows: false, formatColumns: false, insertLink: false, selectCells: true });
            spreadsheet.protectSheet(3, { formatCells: false, formatRows: false, formatColumns: false, insertLink: false, selectCells: true });
            spreadsheet.protectSheet(0, { formatCells: false, formatRows: false, formatColumns: false, insertLink: false, selectCells: false });
            setTimeout(() => {
              spreadsheet.resize();
            }, 50);
          });
        });
    },50);

    //initiating chart models
    const payrollCostByDeptChart: ChartModel[] = [{ type: 'Bar', range: 'Data!A3:B7', title: 'PAYROLL COST BY DEPARTMENT', theme: 'Tailwind3', left: 30, top: 120, width: 475, height: 292, isSeriesInRows: false, primaryYAxis: { visible: false }, dataLabelSettings: { visible: true, position: 'Middle' }, legendSettings: { position: 'Right' } }];
    const overtimeDistributionChart: ChartModel[] = [{ type: 'Doughnut', range: 'Data!C4:D6', title: 'OVERTIME DISTRIBUTION', theme: 'Tailwind3', left: 540, top: 120, width: 475, height: 292, isSeriesInRows: false, dataLabelSettings: { visible: true, position: 'Middle' }, legendSettings: { position: 'Right' } }];
    const leaveDistributionChart: ChartModel[] = [{ type: 'Doughnut', range: 'Data!A15:B17', title: 'LEAVE DISTRIBUTION', theme: 'Tailwind3', left: 1560, top: 120, width: 475, height: 292, isSeriesInRows: false, dataLabelSettings: { visible: true, position: 'Middle' }, legendSettings: { position: 'Right' } }];
    const averageWorkHourChart: ChartModel[] = [{ type: 'Column', range: 'Data!C9:D12', title: 'AVG WORK HOURS/DAY', theme: 'Tailwind3', left: 1050, top: 120, width: 475, height: 292, isSeriesInRows: false, dataLabelSettings: { visible: true, position: 'Outer' }, legendSettings: { position: 'Right' } }];
    const departmentwiseRiskChart: ChartModel[] = [{ type: 'Pie', range: 'Data!A19:B23', title: 'AVG RISK SCORE BY DEPARTMENT', theme: 'Tailwind3', left: 1430, top: 605, width: 600, height: 255, isSeriesInRows: false, dataLabelSettings: { visible: true, position: 'Middle' }, legendSettings: { position: 'Right' } }];
    setCell(0, 0, dashboardSheet, { chart: payrollCostByDeptChart, image: workingHrsImage });
    setCell(0, 1, dashboardSheet, { chart: overtimeDistributionChart, image: dollarImage });
    setCell(0, 2, dashboardSheet, { chart: leaveDistributionChart, image: overtimeHrsImage });
    setCell(0, 3, dashboardSheet, { chart: averageWorkHourChart, image: overtimePayImage });
    setCell(20, 10, dashboardSheet, { chart: departmentwiseRiskChart });
    setCell(0, 4, dashboardSheet, { image: averageWorkImage });
    setCell(0, 5, dashboardSheet, { image: attendanceRateImage });
    setCell(0, 6, dashboardSheet, { image: totalEmployeeImage });
    spreadsheet.resize();
  }

  return (
  <div className='container'><SpreadsheetComponent
    openUrl='https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/open'
    saveUrl='https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/save'
    ref={(ssObj: Spreadsheet) => { spreadsheet = ssObj }} created={onCreated.bind(this)}
    dataSourceChanged={dataSourceChanged}
    cellStyle={{ textAlign: 'center', verticalAlign: 'middle', backgroundColor:'#F5F7FB' }} //Green - #F0FDF4, lavender - #F7F5FF , grey- #F5F7FB
  >
  <SheetsDirective>
    <SheetDirective name='Dashboard' showGridLines={false}></SheetDirective>
    <SheetDirective name='Employee Master' showGridLines={false}></SheetDirective>
    <SheetDirective name='Timesheet' showGridLines={false}></SheetDirective>
    <SheetDirective name='Payroll' showGridLines={false}></SheetDirective>
    <SheetDirective name='Data' state='Hidden'></SheetDirective>
  </SheetsDirective>
  </SpreadsheetComponent></div>);
}