const XLSX = require("xlsx");
const fs = require("fs");

// Read the Excel file
const workbook = XLSX.readFile("data/data.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
// console.log(worksheet)
// Convert the worksheet to an array of objects
const records = XLSX.utils.sheet_to_json(worksheet);
// console.log(records)
// Function to convert Excel serial date to Date object
const excelSerialToDate = (serial) => {
  const baseDate = new Date("1899-12-30T00:00:00Z");
  const millisecondsPerDay = 24 * 60 * 60 * 1000;
  return new Date(baseDate.getTime() + serial * millisecondsPerDay);
};


const employee_with_seven_con_days = [];

const Employee_that_Worked_less_than_10_hours = [];

const Employees_That_Worked_More_Than_14_Hours = [];
// Function to check if an employee has worked for 7 consecutive days
const Employees_That_Worked_Seven_Days = () => {
  const employees = {};
  records.forEach((record) => {
    const employeeName = record["Employee Name"];
    const date = record["Time Out"]
      ? excelSerialToDate(record["Time Out"]).toISOString().split("T")[0]
      : null;
    if (!employees[employeeName]) {
      employees[employeeName] = new Set();
    }
    if (date) {
      employees[employeeName].add(date);
    }
  });

  for (const employeeName in employees) {
    if (employees[employeeName].size >= 7) {
      //   console.log(
      //     `Employee ${employeeName} has worked for 7 consecutive days.`
      //   );
    const newEmployeeName = [];
    newEmployeeName.push(employeeName.split(", "));
    // console.log(newEmployeeName);
    newEmployeeName.forEach((name) => {
      name.forEach((newName) => {
        // Check for duplicates before pushing
        const isDuplicate = new Set(
          employee_with_seven_con_days.map((item) => item.name_of_employee)
        ).has(newName);
        if (!isDuplicate) {
          employee_with_seven_con_days.push({ name_of_employee: newName });
        } 
        
      });
    });
    //   employee_with_seven_con_days.push({ name_of_employee: employeeName });
    }
  }
};

// Function to find employees with less than 10 hours between shifts but greater than 1 hour
const Employees_Less_Than_10_Hours_Between_Shifts = () => {
  const employees = {};
  records.forEach((record, index) => {
    const employeeName = record["Employee Name"];
    if (!employees[employeeName]) {
      employees[employeeName] = [];
    }
    if (index > 0) {
      const previousTimeOut = excelSerialToDate(records[index - 1]["Time Out"]);
      const currentTimeIn = excelSerialToDate(record["Time"]);
      const hoursBetweenShifts =
        (currentTimeIn - previousTimeOut) / (1000 * 60 * 60);
      if (hoursBetweenShifts < 10 && hoursBetweenShifts > 1) {
        // console.log(
        //   `Employee ${employeeName} has less than 10 hours between shifts.`
        // );
       const newEmployeeName = [];
       newEmployeeName.push(employeeName.split(", "));
    //    console.log(newEmployeeName);
       newEmployeeName.forEach((name) => {
         name.forEach((newName) => {
             const isDuplicate = new Set(
               Employee_that_Worked_less_than_10_hours.map(
                 (item) => item.name_of_employee
               )
             ).has(newName);
             if (!isDuplicate) {
            Employee_that_Worked_less_than_10_hours.push({
              name_of_employee: newName,
            });
             } 
          
         });
       });
    //  Employee_that_Worked_less_than_10_hours.push({name_of_employee: employeeName});
      }
    }
    employees[employeeName].push(record);
  });
};

// Function to find employees who have worked for more than 14 hours in a single shift
const Employees_Worked_More_Than_14_Hours = () => {
  records.forEach((record) => {
    const employeeName = record["Employee Name"];
    const timeIn = excelSerialToDate(record["Time"]);
    const timeOut = excelSerialToDate(record["Time Out"]);
    const hoursWorked = (timeOut - timeIn) / (1000 * 60 * 60);
    if (hoursWorked > 14) {
    //   console.log(
    //     `Employee ${employeeName} has worked for more than 14 hours in a single shift.`
    //   );
    const newEmployeeName = [];
    newEmployeeName.push(employeeName.split(", "));
    // console.log(newEmployeeName);
    newEmployeeName.forEach((name) => {
        name.forEach((newName)=>{
         
           const isDuplicate = new Set(
             Employees_That_Worked_More_Than_14_Hours.map(
               (item) => item.name_of_employee
             )
           ).has(newName);
           if (!isDuplicate) {
               Employees_That_Worked_More_Than_14_Hours.push({
                 name_of_employee: newName,
               });
           } 
        })
    
      
    });
    }
  });
};

// Call for the functions
var colors = require("colors");
Employees_That_Worked_Seven_Days();
console.log("Employee That Worked For Consecutive Seven Days".cyan);
console.log(employee_with_seven_con_days)
Employees_Less_Than_10_Hours_Between_Shifts();
console.log("Employee That Worked Less Than 10 Hours and More than 1 Hour".red);
console.log(Employee_that_Worked_less_than_10_hours)
Employees_Worked_More_Than_14_Hours();
console.log("Employee That Worked More Than 14 Hours in a Single Shift".yellow);
console.log(Employees_That_Worked_More_Than_14_Hours);
