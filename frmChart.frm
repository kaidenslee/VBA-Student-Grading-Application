VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChart 
   Caption         =   "Create Chart"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6540
   OleObjectBlob   =   "frmChart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =========================================================+
' Date: April 3rd, 2025
' Program title: Chart Creation Userform
' Description: The interface for creating a chart
'===========================================================+
'instructions
Private Sub UserForm_Initialize()
    MsgBox "To use this tool, select an assignment and press 'Create Chart' to create a graphical chart for that assignment."
End Sub
'if user selects back to main button, hide this form and display the main form again
Private Sub btnBackToMain_Click()
    Me.Hide
    frmGrades.Show
End Sub

'going to create a chart based on options selected
Private Sub btnCreateChart_Click()
    Dim selectedField As String
    Dim dbPath As String
    Dim conn As Object, rs As Object
    Dim sql As String
    Dim ws As Worksheet
    Dim i As Integer
    Dim chartObj As ChartObject

    ' find selected option button
    Select Case True
        Case optA1.Value: selectedField = "A1"
        Case optA2.Value: selectedField = "A2"
        Case optA3.Value: selectedField = "A3"
        Case optA4.Value: selectedField = "A4"
        Case optMidterm.Value: selectedField = "Midterm"
        Case optFinalExam.Value: selectedField = "[Final Exam]"
        Case Else
        'in case no options were selected
            MsgBox "Please select a grade item to chart.", vbExclamation
            Exit Sub
    End Select

'set connection to database
    dbPath = frmGrades.txtFilePath.Text
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

    sql = "SELECT " & selectedField & " FROM Grades WHERE " & selectedField & " IS NOT NULL"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1

    ' output to new worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = selectedField & "_Chart"
    ws.Cells(1, 1).Value = selectedField

    i = 2
    Do While Not rs.EOF
        ws.Cells(i, 1).Value = rs.Fields(0).Value
        i = i + 1
        rs.MoveNext
    Loop

    rs.Close
    conn.Close

    ' insert chart
    Set chartObj = ws.ChartObjects.Add(Left:=250, Width:=400, Top:=10, Height:=300)

    With chartObj.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ws.Range("A2:A" & i - 1)
        .HasTitle = True
        .ChartTitle.Text = selectedField & " Grade Distribution"
        .Axes(xlCategory).HasTitle = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Grade"
    End With

'notify user of successful creation
    MsgBox selectedField & " chart created!", vbInformation
End Sub
