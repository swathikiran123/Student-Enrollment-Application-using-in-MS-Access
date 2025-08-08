# Student-Enrollment-Application-using-in-MS-Access
This project is a Student Enrollment Management System built using Microsoft Access (2007-2016 format) with VBA (Visual Basic for Applications) to streamline and manage student data for academic institutions.

Project Overview
This MS Access application enables users to:

Store student details (name, gender, DOB, contact, course, etc.)
Enter and update student records using a user-friendly form
Auto-generate Student IDs
Generate and view student reports with a click of a button
Validate data input using VBA logic
Provide a lightweight and offline database solution
Database Structure
Table: Students

Field Name	Data Type	Description
StudentID	AutoNumber	Primary Key, Auto-generated
FullName	Short Text	Student's full name
Gender	Short Text	Gender (Male/Female)
DOB	Date/Time	Date of Birth
PhoneNumber	Short Text	Contact Number
Email	Short Text	Email Address
Course	Short Text	Enrolled Course
EnrolmentDate	Date/Time	Date of Enrollment
Forms
🔹 Std_Entry_Form
A form to add, edit, and view student data.
Includes text boxes and combo boxes for all student fields.
Uses VBA event handlers to:
Validate inputs
Submit records to the table
Trigger reports
Reports
🔹 rpt_StudentDetails
A printable report of all student records.
Launched via a button on the form using a VBA macro.
Private Sub form_load()
  MsgBox "Welcome to data entry form"  
End Sub
