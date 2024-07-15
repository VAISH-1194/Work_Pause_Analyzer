# Work_Pause_Analyzer

**Work Pause Analyzer** is designed to streamline the process of managing employee work and break times efficiently. This user-friendly platform offers two primary upload options, each tailored to simplify specific tasks:
**1.	File Splitting:** Users can upload large files and split them into manageable segments, facilitating organized and accessible data management.
**2.	Break Time Calculation:** This feature extends beyond basic time tracking by calculating:
->	**Intime:** The time an employee starts their workday.
->	**Outtime:** The time an employee ends their workday.
->	**Total Duration:** The total working hours each day.
->	**Leave Counts:** The number of leave days taken.
->	**Break Time:** The total break time availed during the workday.
To enhance user comfort, the platform includes both dark and light mode options, allowing users to switch based on their preference and reducing eye strain.
Work Pause Analyzer aims to help businesses optimize their workforce management through its powerful yet easy-to-use features, making it an essential tool for enhancing organizational productivity and efficiency.

**Project Overview: Work Pause Analyzer**

**Introduction:** Work Pause Analyzer is a sophisticated web application designed to optimize the management of employee work and break times. It offers a user-friendly interface and powerful features to help businesses streamline their workforce management.
**Key Features:**
**1.	File Splitting:**
->	**Motivation for File Splitting:** During the development and initial testing phases, we encountered significant challenges in processing large tables with complex data arrangements. These large files often led to inefficiencies, slow processing times, and occasional errors, making it difficult to manage and analyze the data effectively.
->	**Solution:** To address these issues, we implemented a file splitting feature. This allows users to upload large files and split them into more manageable segments. By breaking down the data into smaller parts, we can ensure smoother processing, improved data integrity, and enhanced performance.


**In Time**

The "In Time" is defined as the time when an employee starts their workday. To determine this accurately, the system processes the punch records as follows:

-> 	Retrieval of Punch Records
->	Identifying the First Punch
->	Handling Multiple Punches

**Out Time**

The "Out Time" is defined as the time when an employee ends their workday. The process to determine "Out Time" is as follows:

->	**Identifying the Last Punch:**
The system identifies the last timestamp in the "Punch Records" column.
This timestamp represents the last instance the employee punches out for the day, which is considered the "Out Time."
->	**Handling Missing Out Records:**
In cases where the last punch record does not indicate an "out" entry (e.g., due to missing records or incomplete data), the system appends the term "records missing" to the last entry.
This alerts the user to potential issues with the recorded data and ensures transparency in the time tracking process.

**Challenges encountered**
During the processing of punch records, we faced significant challenges due to:
->	**Invalid Entries:** Entries that do not correspond to valid punch-in or punch-out actions.
->	**Missing Entries:** Incomplete records that lack either the punch-in or punch-out times.
Filtering out only the valid entries often led to inaccuracies in calculating the total duration of work, as missing entries resulted in incomplete data.

**Steps involved in the Processing**

**Read the Excel File:**
•	The application reads the uploaded Excel file containing the punch records.
  Drop Rows Starting with "Total":
•	Any row that starts with "Total" is identified and removed to ensure the data set contains only relevant entries.
  Rearrange Columns:
•	The rows labeled "Department" and "Emp code" are moved to the end of the table for better organization.
 
**Set Header and Clean Columns:**
•	The header row is set to the row labeled "Att. Date," and any columns with unnamed headers are dropped.
  **Update "InTime" and "OutTime":**
•	The "InTime" is determined by the first entry in the "Punch Records" column.
•	The "OutTime" is determined by the last entry in the "Punch Records" column.
•	If the last entry is not an "out" entry, it is updated to indicate "records missing."
  Remove "1st" Floor Entries:
•	Entries from the 1st floor, which are not used for break time calculation, are removed from the data set.
  **Remove Invalid Entries:**
•	Any entries that do not contain valid "in" or "out" times are identified and removed to ensure data accuracy.
  **Create a Copy of Corrected Records:**
•	A copy of the corrected records is made for display purposes.
  **Calculate Total Duration:**
•	The total duration of work for each employee is calculated using the "InTime" and "OutTime."
  **Update Employee Status:**
•	The status of each employee is updated to indicate whether they are "Present" or "Absent" based on their punch records.
  **Update Record Status:**
•	The status of each record is updated to indicate whether it is "Valid" or "Invalid."
**Calculate Approximate Break Time:**
•	The approximate break time is calculated by adding a duplicate entry labeled "--:--:(ED)" to indicate the end of the day.
 **Convert Break Time to Hours and Minutes:**
•	The break time, initially calculated in minutes, is converted to hours and minutes for easier interpretation.
 **Drop Unwanted Columns:**
•	Any columns that are not necessary for the final report are removed from the data set.
  Calculate Leave Dates and Number of Leaves:
•	The leave dates are identified, and the total number of leave days is calculated for each employee.
  **Style and Adjust Cell Sizes:**
•	The resultant Excel file is styled and the cell sizes are adjusted for better readability and presentation.
  **Process and Download Files:**
•	The processed files are compiled into a zip file, which can be downloaded by the user. This zip file contains all the processed records in a structured format.
These steps ensure that the data is meticulously processed and accurately reflects the work and break times of employees. By following this comprehensive process, Work Pause Analyzer provides a reliable and user-friendly solution for workforce management.
**Benefits:**
->	Enhanced Productivity: By providing detailed insights into work patterns, businesses can identify and address areas for improvement.
->	User-Friendly Interface: Designed to be accessible for users of all skill levels.
->	Efficient Data Management: The file splitting feature aids in handling extensive datasets effectively.
Work Pause Analyzer aims to help businesses optimize their workforce management through its powerful yet easy-to-use features, making it an essential tool for enhancing organizational productivity and efficiency.

![image](https://github.com/user-attachments/assets/460e7543-5ec3-4f31-9851-5132b5569793)

![Screenshot 2024-07-15 024559](https://github.com/user-attachments/assets/c4ce9510-a7b4-4514-ac33-4dc2060a5cfc)

![Screenshot 2024-07-15 021321](https://github.com/user-attachments/assets/28c97afb-253f-4757-ab17-ea454aa2d0d9)

![image](https://github.com/user-attachments/assets/0f9ad7ee-093f-4399-a6b4-fe1b247cf704)




