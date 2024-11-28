# Text String Manipulation and Data Aggregation in Excel
## Project Objective
To demonstrate Excel skills by cleaning and manipulating text data within an HR dataset, this project highlights the use of Excel functions to handle common data cleaning tasks such as splitting, joining, and validating data fields. Additionally, the file includes an "HR Analytics" sheet showcasing aggregation functions in Excel.

## Dataset(s)
- The <a href="https://github.com/DennyMandaka/Text-String-Manipulation-and-Data-Aggregation-in-Excel/blob/main/Text%20String%20Manipulation%20and%20Data%20Aggregation%20in%20Excel.xlsx">dataset</a> was sourced from the "Excel Skills for Data Analytics and Visualization" course offered by Macquarie University on Coursera. It contains information from an HR database, including employee IDs, names, emails, hire dates, departments, office locations, hourly rates, leave days, and more.

## Data Cleaning and Text String Manipulation
- In C4, in the HR Data sheet join the data in A4 and B4 to create the employee ID F1180. Copy the formula down for the rest of the staff.
- In K4, extract just the floor number from J4 (first 2 numbers) and then copy it down.
- In F4, generate the full name for each staff member by joining the first name, a space and the surname.
- In G4, generate an email address by joining the first letter of the first name, the surname and "@zenco.com" to get "Sbacata@zenco.com".
- In L4, extract the text between the two dashes (West) in Location then copy down to make sure it works for all staff.
- In M4, extract the last 4 characters from the Location in J4.

## Aggregating and Summarizing Data
- Use the Name Box to change the name of cell G2 to Overtime.
- In G5 calculate the overtime rate by multiplying the Hourly Rate by Overtime. Copy the formula down.
- Use Create from Selection to name each of the columns of employee data.
- In J2 calculate the total number of Days Sick.
- Open the Name Manager and create a new constant called Leave_Allowance with the value 20.
- In J5 calculate how much leave is still available by subtracting Leave Taken from Leave Allowance. 
- Use Create from Selection to name the list of departments in column M.
- Go to the Name Box and select Department to highlight the staff department data.
- Apply Data Validation to the column so that each cell gets a drop-down list of valid departments. (Make sure to use Departments not Department)
- In N5 calculate the total leave available in the Accounting Department. 
- In O5 calculate the average Hourly Rate for staff who are Part Time and in the Accounting Department (i.e. 2 criteria M5 and O4).
- Apply appropriate cell referencing to M5 and O4 so that the formula can be dragged down and then across.

## Process
- Use & and CONCAT to join text strings.
- Use PROPER and LOWER to change text cases.
- Use LEFT, MID, and RIGHT to extract parts of text strings.
- Use CLEAN and TRIM to remove extra spaces and other non-printing characters.
- Use SUBSTITUTE, CODE, and UNICHAR to remove other non-printing characters that the CLEAN function cannot remove.
- Use the Data Validation to create a drop-down list.
- Use basic aggregating function like SUM, SUMIFS, and AVERAGEIFS.
- The <a href="https://github.com/DennyMandaka/Text-String-Manipulation-and-Data-Aggregation-in-Excel/blob/main/Data%20Cleaning%20and%20Text%20String%20Manipulation.png">result</a> is like this and <a href="https://github.com/DennyMandaka/Text-String-Manipulation-and-Data-Aggregation-in-Excel/blob/main/Data%20Aggregation.png">like this</a>.

## Insights and Conclusion
- Utilized Excel's functions to demonstrate data cleaning and text string manipulation.
- The data cleaning process highlighted the importance of consistent entry standards.
- Demonstrated the use of some of Excel’s most commonly used aggregation functions to perform calculations and analysis.
