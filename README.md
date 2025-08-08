# Student-Enrollment-Application-using-in-MS-Access
This project is a Student Enrollment Management System built using Microsoft Access (2007-2016 format) with VBA (Visual Basic for Applications) to streamline and manage student data for academic institutions.

**Project Overview**

This MS Access application enables users to:


Store student details (name, gender, DOB, contact, course, etc.)
Enter and update student records using a user-friendly form
Auto-generate Student IDs
Generate and view student reports with a click of a button
Validate data input using VBA logic
Provide a lightweight and offline database solution

**Database Structure**

Table: Students

| Field Name      | Data Type  | Description                     |
| --------------- | ---------- | ------------------------------- |
| `StudentID`     | AutoNumber | **Primary Key**, auto-generated |
| `FullName`      | Short Text | Full name of the student        |
| `Gender`        | Short Text | Gender (Male/Female)            |
| `DOB`           | Date/Time  | Date of Birth                   |
| `PhoneNumber`   | Short Text | Contact number                  |
| `Email`         | Short Text | Email address                   |
| `Course`        | Short Text | Enrolled course name            |
| `EnrolmentDate` | Date/Time  | Date of enrollment              |


**Forms**

🔹 Std_Entry_Form

A form to add, edit, and view student data.
Includes text boxes and combo boxes for all student fields.
Uses VBA event handlers to:
Validate inputs
Submit records to the table
Trigger reports

**Reports**

🔹 rpt_StudentDetails

A printable report of all student records.
Launched via a button on the form using a VBA macro.

    
Private Sub Form_Load

    MsgBox "Welcome to the Data Entry Form"
    
End Sub

