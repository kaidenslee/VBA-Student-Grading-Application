VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmViewGrades 
   Caption         =   "View Grades by Student"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmViewGrades.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmViewGrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==========================================================+
' Date: April 3rd, 2025
' Program title: View Grades Userform
' Description: The interface for viewing grades by course and displaying
'===========================================================+
Private Sub btnViewGrades_Click()
    Dim selectedCourse As String
    Dim ctrl As control
    Dim dbPath As String
    Dim conn As Object, rs As Object
    Dim sql As String
    Dim ws As Worksheet
    Dim row As Long, col As Long

    'get selected course code
    selectedCourse = ""

    For Each ctrl In fraCourses.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.Value = True Then
                selectedCourse = ctrl.Tag
                Exit For
            End If
        End If
    Next ctrl

    If selectedCourse = "" Then
        MsgBox "Please select a course!", vbExclamation
        Exit Sub
    End If

    'connect to database
    dbPath = frmGrades.txtFilePath.Text
    If dbPath = "" Then
        MsgBox "No database path found!", vbCritical
        Exit Sub
    End If

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

    'build sql
    If selectedCourse = "ALL" Then
        sql = "SELECT Grades.StudentID, FirstName, LastName, Course, A1, A2, A3, A4, MidTerm, [Final Exam] " & _
              "FROM Grades INNER JOIN Students ON Grades.StudentID = Students.StudentID"
    Else
        sql = "SELECT Grades.StudentID, FirstName, LastName, Course, A1, A2, A3, A4, MidTerm, [Final Exam] " & _
              "FROM Grades INNER JOIN Students ON Grades.StudentID = Students.StudentID " & _
              "WHERE Course = '" & selectedCourse & "'"
    End If

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1

    If rs.EOF Then
        MsgBox "No students found for that course!", vbInformation
        rs.Close: conn.Close
        Exit Sub
    End If

    'write to Excel
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Grades_" & IIf(selectedCourse = "ALL", "All", selectedCourse)

    'create heeaders
    For col = 0 To rs.Fields.count - 1
        ws.Cells(1, col + 1).Value = rs.Fields(col).Name
    Next col

    'data
    row = 2
    Do While Not rs.EOF
        For col = 0 To rs.Fields.count - 1
            ws.Cells(row, col + 1).Value = rs.Fields(col).Value
        Next col
        row = row + 1
        rs.MoveNext
    Loop

    rs.Close: conn.Close
'notify user of success
    MsgBox "Grades for " & selectedCourse & " exported to Excel!", vbInformation
End Sub
'copied from previous 'class average userform code, to get dynamically created buttons based on course code. comments there
Private Sub UserForm_Initialize()
'instructions
    MsgBox "To use this tool, select a course option, and press continue to display the student info for that course."
    On Error GoTo LoadError

    Dim conn As Object, rs As Object
    Dim dbPath As String, connString As String, sql As String
    Dim optBtn As MSForms.OptionButton
    Dim topOffset As Integer: topOffset = 10

    dbPath = frmGrades.txtFilePath.Text
    If dbPath = "" Then
        MsgBox "No database path provided.", vbExclamation
        Unload Me
        Exit Sub
    End If

    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connString

    sql = "SELECT CourseCode, CourseName FROM Courses"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1

    Do While Not rs.EOF
        Set optBtn = fraCourses.Controls.Add("Forms.OptionButton.1", , True)
        With optBtn
            .Caption = rs("CourseName")
            .Tag = rs("CourseCode")
            .Top = topOffset
            .Left = 10
            .Width = 200
        End With
        topOffset = topOffset + 20
        rs.MoveNext
    Loop

    Set optBtn = fraCourses.Controls.Add("Forms.OptionButton.1", , True)
    With optBtn
        .Caption = "All Courses"
        .Tag = "ALL"
        .Top = topOffset
        .Left = 10
        .Width = 200
    End With

    rs.Close
    conn.Close
    Exit Sub

LoadError:
    MsgBox "Error loading courses: " & Err.Description, vbCritical
    Unload Me
End Sub

Private Sub btnBackToMain_Click()
    Me.Hide
    frmGrades.Show
End Sub
