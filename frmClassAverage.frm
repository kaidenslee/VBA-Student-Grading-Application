VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClassAverage 
   Caption         =   "Calculate Class Average"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmClassAverage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClassAverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =======================================================+
' Date: April 3rd, 2025
' Program title: Class average userform
' Description: The interface for calculating a class average
'===========================================================+
' instructions
Private Sub UserForm_Initialize()
    MsgBox "To use this tool, select a course option and press continue to see the average for that course."
    On Error GoTo LoadError

    Dim conn As Object, rs As Object
    Dim dbPath As String, connString As String, sql As String
    Dim optBtn As MSForms.OptionButton
    Dim topOffset As Integer: topOffset = 10
'set database connection
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
'create option buttons for all the courses
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

    ' add "All Courses" as option
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
'if user presses return to main button, close this form and return to main form
Private Sub btnBackToMain_Click()
    Me.Hide
    frmGrades.Show
End Sub

Private Sub btnCalcAverage_Click()
    Dim selectedCourseID As String
    Dim ctrl As control
    Dim conn As Object, rs As Object
    Dim dbPath As String, sql As String
    Dim grades() As Double
    Dim total As Double, gradeCount As Long, i As Long
    Dim avg As Double, stdev As Double, sumSqDiff As Double

    'get selected course
    For Each ctrl In fraCourses.Controls
        If TypeName(ctrl) = "OptionButton" And ctrl.Value = True Then
            selectedCourseID = ctrl.Tag
            Exit For
        End If
    Next ctrl

    If selectedCourseID = "" Then
        MsgBox "Select a course!", vbExclamation
        Exit Sub
    End If

    'connect to database
    dbPath = frmGrades.txtFilePath.Text
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

    ' build SQL
    If selectedCourseID = "ALL" Then
        sql = "SELECT (A1*0.05 + A2*0.05 + A3*0.05 + A4*0.05 + MidTerm*0.3 + [Final Exam]*0.5) AS FinalGrade FROM Grades"
    Else
        sql = "SELECT (A1*0.05 + A2*0.05 + A3*0.05 + A4*0.05 + MidTerm*0.3 + [Final Exam]*0.5) AS FinalGrade " & _
              "FROM Grades WHERE Course = '" & selectedCourseID & "'"
    End If

    Set rs = conn.Execute(sql)

    total = 0: gradeCount = 0
    Do While Not rs.EOF
        ReDim Preserve grades(gradeCount)
        grades(gradeCount) = rs("FinalGrade")
        total = total + grades(gradeCount)
        gradeCount = gradeCount + 1
        rs.MoveNext
    Loop

    If gradeCount = 0 Then
        MsgBox "No grades found!", vbInformation
        rs.Close: conn.Close
        Exit Sub
    End If

    avg = total / gradeCount
    sumSqDiff = 0

    For i = 0 To gradeCount - 1
        sumSqDiff = sumSqDiff + (grades(i) - avg) ^ 2
    Next i

    stdev = Sqr(sumSqDiff / gradeCount)

    rs.Close
    conn.Close

    MsgBox "Final Grade Stats:" & vbCrLf & _
           "Average: " & Format(avg, "0.00") & "%" & vbCrLf & _
           "Standard Deviation: " & Format(stdev, "0.00") & "%", _
           vbInformation, "Class Summary"
End Sub

