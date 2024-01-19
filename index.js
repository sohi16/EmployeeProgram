const path = require('path');
const XLSX = require('xlsx');
const fs = require('fs');

// Load the spreadsheet
const file_path = path.join(__dirname, 'Assignment_Timecard.xlsx');
const workbook = XLSX.readFile(file_path);
const sheet_name = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheet_name];

// Convert the sheet to JSON
const data = XLSX.utils.sheet_to_json(worksheet);

// Sort data by employee and date
data.sort((a, b) => {
    const employeeComparison = (a['Employee Name'] || '').localeCompare(b['Employee Name'] || '');
    if (employeeComparison !== 0) {
        return employeeComparison;
    }

    return new Date(a['Pay Cycle Start Date'] || '') - new Date(b['Pay Cycle Start Date'] || '');
});

// Function to check if an employee has worked for 7 consecutive days
function hasConsecutiveDays(employeeData) {
    for (let i = 1; i < employeeData.length; i++) {
        const prevDate = new Date(employeeData[i - 1]['Pay Cycle Start Date']);
        const currentDate = new Date(employeeData[i]['Pay Cycle Start Date']);
        const timeDiff = currentDate.getTime() - prevDate.getTime();
        const daysDiff = timeDiff / (1000 * 3600 * 24);

        if (daysDiff !== 1) {
            return false;
        }
    }
    return true;
}

// Analyze data and generate output
let output = '';
let currentEmployee = null;
let currentEmployeeData = [];

data.forEach(entry => {
    if (currentEmployee !== entry['Employee Name']) {
        if (hasConsecutiveDays(currentEmployeeData)) {
            output += `${currentEmployee} has worked for 7 consecutive days.\n`;
        }

        // Reset for the next employee
        currentEmployee = entry['Employee Name'];
        currentEmployeeData = [entry];
    } else {
        currentEmployeeData.push(entry);
    }
});

// Check for consecutive days for the last employee
if (hasConsecutiveDays(currentEmployeeData)) {
    output += `${currentEmployee} has worked for 7 consecutive days.\n`;
}

// Save output to a file
fs.writeFileSync('output.txt', output);

// Print the output to the console
console.log(output);
