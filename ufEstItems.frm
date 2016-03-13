VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufEstItems 
   Caption         =   "Estimate Item Builder"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10350
   OleObjectBlob   =   "ufEstItems.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufEstItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private xmlhand As New XMLHandler
Private utils As New Utilities

Private Sub btnAddVersion_Click()
    Dim intChoice As Integer
    Dim strPath As String
    
    'only allow the user to select one file
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    'Adjust title of dialog
    Application.FileDialog(msoFileDialogOpen).Title = "Please select an estimate graphic"
    'make the file dialog visible to the user
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    'determine what choice the user made
    If intChoice <> 0 Then
        'get the file path selected by the user
        strPath = Application.FileDialog( _
            msoFileDialogOpen).SelectedItems(1)
        'print the file path to sheet 1
        Me.txtVersionPath.Value = utils.FileNameOnly(strPath)
    End If
End Sub
Private Sub btnClean_Click()
    Dim cleanName As String
    Dim rng As Range
    Dim c As Range

    Dim ac As Integer
    Dim lastr As Long
    
    ac = ActiveCell.column
    lastr = Sheets("Sheet1").Cells(Rows.Count, ac).End(xlUp).Row
    Set rng = Range(Cells(1, ac), Cells(lastr, ac))
    Application.ScreenUpdating = False

    For Each c In rng
        cleanName = c.Value '.Replace(".", "")
        c.Value = Replace(cleanName, ".", "")
    Next c
    
    Application.ScreenUpdating = True
End Sub
Private Sub btnDefaults_Click()
    ufPreferences.Show
End Sub
Private Sub btnLast_Click()
    ButtonUI.btnLast
End Sub
Private Sub btnNext_Click()
    Dim r As Long
    
    If IsNumeric(ufEstItems.txtEIRowNumber.Text) Then
        r = CLng(ufEstItems.txtEIRowNumber.Text)
        
        r = r + 1
        If r > 1 And r <= LastRow Then
            ufEstItems.txtEIRowNumber.Text = FormatNumber(r, 0)
            
        End If
    End If
End Sub
Private Sub btnPrev_Click()
    ButtonUI.btnPrevious
End Sub
Private Sub btnXML_Click()
    xmlhand.writeXMLNestRequest
End Sub
Private Sub txtEIRowNumber_Change()
    ButtonUI.GetEIData
End Sub
Private Sub UserForm_Activate()
    ButtonUI.GetEIData
End Sub
Private Sub UserForm_Initialize()
   LoadComboBox cbMaterial, ThisWorkbook.Sheets("Substrates").Range("Names")
   ButtonUI.GetEIData
   ButtonUI.LastRow = utils.FindLastRow()
    With Me.cbShapeFactor
        .AddItem "Rectangle"
        .AddItem "Round Rect"
        .AddItem "Oval"
        .AddItem "Star"
        .AddItem "Cut"
    End With
    With Me.cbSides
        .AddItem "1"
        .AddItem "2"
    End With
    With Me.cbRotation
        .AddItem "any"
        .AddItem "0"
        .AddItem "90"
        .AddItem "180"
        .AddItem "270"
    End With
End Sub
Private Sub LoadComboBox(cBox As ComboBox, listRange As Range)
    Dim tmpAry
    tmpAry = listRange
        With cBox
            .Clear
            .ColumnCount = listRange.Columns.Count
            .List = tmpAry
            
        End With
    Erase tmpAry
    'tmpArray = Null
End Sub
Private Sub txtSubst_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'++++++++++++++++++++++++++++++++++++
    
'  Sheets("Sheet1").Range("a1").Value = Me.txtSubst.Text
'++++++++++++++++++++++++++++++++++++
'  oRange.ClearContents
End Sub


