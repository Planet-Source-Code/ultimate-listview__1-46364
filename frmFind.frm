VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ListView Finder"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "frmFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdvQry 
      Caption         =   "&Advanced Query >>"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtQRYHidden 
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdShowSQL 
      Caption         =   "Show S&QL"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Tag             =   "QryCtls"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdRunQuery 
      Caption         =   "&Run Query"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Tag             =   "QryCtls"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCriteria2 
      Height          =   315
      Left            =   6240
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox txtCriteria1 
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Tag             =   "QryCtls"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.ComboBox cmbFld 
      Height          =   315
      ItemData        =   "frmFind.frx":5C12
      Left            =   2520
      List            =   "frmFind.frx":5C14
      TabIndex        =   4
      Tag             =   "QryCtls"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbQryArg 
      Height          =   315
      ItemData        =   "frmFind.frx":5C16
      Left            =   3720
      List            =   "frmFind.frx":5C35
      TabIndex        =   5
      Tag             =   "QryCtls"
      Text            =   "="
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cdlExport 
      Left            =   1080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   375
      Left            =   10440
      TabIndex        =   12
      Top             =   5400
      Width           =   1100
   End
   Begin VB.ComboBox cmbOutputTo 
      Height          =   315
      ItemData        =   "frmFind.frx":5C67
      Left            =   9000
      List            =   "frmFind.frx":5C7D
      TabIndex        =   11
      Text            =   "Screen"
      Top             =   5460
      Width           =   1335
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   10440
      TabIndex        =   13
      Top             =   6240
      Width           =   1100
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5741
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "imglstListImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imglstListImages 
      Left            =   360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":5CAB
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":5E05
            Key             =   "Down"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   2400
      TabIndex        =   18
      Tag             =   "QryCtls"
      Top             =   5280
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdShowAll 
         Caption         =   "Show &All"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   720
         Width           =   1000
      End
      Begin VB.Label lblAnd 
         Caption         =   "and"
         Height          =   315
         Left            =   3460
         TabIndex        =   19
         Top             =   280
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.Label lblOutputTo 
      Caption         =   "Output To:"
      Height          =   195
      Left            =   8160
      TabIndex        =   16
      Top             =   5580
      Width           =   855
   End
   Begin VB.Label lblFldLike 
      AutoSize        =   -1  'True
      Caption         =   "Click a Column Header for Search Function"
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   1170
      Width           =   3060
   End
   Begin VB.Label lblInstruct 
      Caption         =   $"frmFind.frx":5F5F
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   8175
   End
   Begin VB.Label lblRecCount 
      AutoSize        =   -1  'True
      Caption         =   "lblrecCount"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   4920
      Width           =   795
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ExportReport(ExportType As String, rptExport As Object, FileNm As String)
Dim oPDF As ActiveReportsPDFExport.ARExportPDF
Dim oEXL As ActiveReportsExcelExport.ARExportExcel
Dim oTXT As ActiveReportsTextExport.ARExportText
Dim oHTML As ActiveReportsHTMLExport.HTMLexport
'Dim oHTML As HTMLexport
Dim oTIFF As ActiveReportsTIFFExport.TIFFExport

rptExport.Run


Select Case ExportType
    
        Case "PDF"
            Set oPDF = New ActiveReportsPDFExport.ARExportPDF
            oPDF.FileName = App.Path & "\" & FileNm & ".PDF"
            oPDF.Export rptExport.Pages

            
        Case "Excel"
            Set oEXL = New ActiveReportsExcelExport.ARExportExcel
            oEXL.FileName = App.Path & "\" & FileNm & ".xls"
            oEXL.Export rptExport.Pages
            
        Case "Text"
            Set oTXT = New ActiveReportsTextExport.ARExportText
            oTXT.FileName = App.Path & "\" & FileNm & ".txt"
            oTXT.PageDelimiter = ";"
            oTXT.TextDelimiter = ","
            oTXT.Export rptExport.Pages
            

        Case "HTML"
            Set oHTML = New ActiveReportsHTMLExport.HTMLexport
            oHTML.FileNamePrefix = FileNm
            oHTML.HTMLOutputPath = App.Path
            oHTML.Export rptExport.Pages

        
        Case "Tiff"
        MsgBox "tiff"
            Set oTIFF = New ActiveReportsTIFFExport.TIFFExport
            oTIFF.FileName = App.Path & "\" & FileNm & ".tiff"
            oTIFF.Export rptExport.Pages
                
End Select


End Sub

Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)
    
    Dim x As Long, y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For x = 0 To Xupper
        For y = 0 To Yupper
            tempArray(x, y) = v(y, x)
        Next y
    Next x
    
    TransposeDim = tempArray

End Function


Private Sub cmbFld_Click()

cmbQryArg.SetFocus

End Sub

Private Sub cmbQryArg_Click()

If cmbQryArg.Text = "Between" Then
txtCriteria2.Visible = True: lblAnd.Visible = True
Else
txtCriteria2.Visible = False: lblAnd.Visible = False
End If

txtCriteria1.SetFocus


End Sub

Private Sub cmdAdvQry_Click()

Dim ctl As Control
If cmdAdvQry.Caption = "&Advanced Query >>" Then
cmdAdvQry.Caption = "<< &Simple"
For Each ctl In Me.Controls
If ctl.Tag = "QryCtls" Then
ctl.Visible = True
End If
Next ctl

Else
cmdAdvQry.Caption = "&Advanced Query >>"
For Each ctl In Me.Controls
If ctl.Tag = "QryCtls" Then
ctl.Visible = False
End If
Next ctl
lblAnd.Visible = False
txtCriteria2.Visible = False

End If


End Sub

Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdGO_Click()
    
    Dim rpt As New LVTemplate
    Set rpt.Grid = ListView1

Select Case cmbOutputTo

Case "Screen"
    rpt.Show


Case "HTML"

Screen.MousePointer = vbHourglass
    rpt.Run
    ExportReport "HTML", rpt, TblName
    Screen.MousePointer = vbDefault
    MsgBox "HTML saved to " & App.Path & "\" & TblName & ".htm"

Case "PDF"

Screen.MousePointer = vbHourglass
    rpt.Run
    ExportReport "PDF", rpt, TblName
Screen.MousePointer = vbDefault
MsgBox "PDF saved to " & App.Path & "\" & TblName & ".pdf"


Case "Printer"
MsgBox "Printer"

Case "CSV"
  On Error GoTo ErrHand
  cdlExport.CancelError = True
  cdlExport.Filter = "csv (*.csv)|*.csv|Text (*.txt)|*.txt"
  cdlExport.FilterIndex = 1
  cdlExport.FileName = ""
  cdlExport.DefaultExt = "csv"
  cdlExport.DialogTitle = "Save As"
 'cdlExport.hWnd = Me.hWnd
  cdlExport.InitDir = App.Path
  cdlExport.Flags = cdlOFNOverwritePrompt
  cdlExport.ShowSave

ExportCSV cdlExport.FileTitle

GoSub AskViewExcel

Case "Excel"
  On Error GoTo ErrHand
  cdlExport.CancelError = True
  cdlExport.Filter = "Excel (*.xls)|*.xls|Any File (*.*)|*.*"
  cdlExport.FilterIndex = 1
  cdlExport.FileName = ""
  cdlExport.DefaultExt = "xls"
  cdlExport.DialogTitle = "Save As"
 'cdlExport.hWnd = Me.hWnd
  cdlExport.InitDir = App.Path
  cdlExport.Flags = cdlOFNOverwritePrompt
  cdlExport.ShowSave

ExportToExcel ListView1, cdlExport.FileTitle

GoSub AskViewExcel


Case "Clipboard"

SendToClipboard ListView1

Case Else
Exit Sub

End Select

Exit Sub


ErrHand:
Err.Clear

Exit Sub

AskViewExcel:

Dim MyMsg
MyMsg = MsgBox("Do you want to view the results in Excel now?", vbQuestion + _
    vbYesNo + vbDefaultButton1, "View now")

If MyMsg = vbYes Then

'Open the new text file in Excel
   Shell "C:\Program Files\Microsoft Office\Office\Excel.exe " & _
      Chr(34) & cdlExport.FileTitle & Chr(34), vbMaximizedFocus
End If
'Return

End Sub


Private Sub Combo1_Change()

End Sub

Private Sub cmdRunQuery_Click()


'On Error GoTo ErrHand

Dim myQry As String

If cmbQryArg = "Like" Or cmbQryArg = "Not Like" Then
txtCriteria1 = "%" & txtCriteria1 & "%"
End If


If IsNumeric(txtCriteria1) = False Then
txtCriteria1 = "'" & txtCriteria1 & "'"
End If
If IsNumeric(txtCriteria2) = False Then
txtCriteria2 = "'" & txtCriteria2 & "'"
End If

RunTheQuery:
If txtCriteria2.Visible = False Then
myQry = "SELECT * FROM " & TblName & " WHERE " & "[" & cmbFld & "]" & " " & cmbQryArg & " " & txtCriteria1
Else
myQry = "SELECT * FROM " & TblName & " WHERE " & "[" & cmbFld & "]" & " " & cmbQryArg & " " & txtCriteria1 & " and " & txtCriteria2
End If

txtQRYHidden = myQry

txtCriteria1 = "": txtCriteria2 = ""

If oRs.State = 1 Then
oRs.Close
Set oRs = Nothing
End If
If oConn.State = 1 Then
oConn.Close
End If

FillListView myQry, ListView1


If ListView1.ListItems.Count > 0 Then
cmdShowAll.Enabled = True
lblRecCount = ListView1.ListItems.Count & _
" Records found in table : " & TblName
Else
lblRecCount = ListView1.ListItems.Count & _
" Record found in table : " & TblName
End If

Exit Sub

ErrHand:

MsgBox "No Data to show, Try a diffeent selection"
Err.Clear

End Sub

Private Sub cmdShowAll_Click()

If oRs.State = 1 Then
oRs.Close
Set oRs = Nothing
End If
If oConn.State = 1 Then oConn.Close

Form_Load
End Sub

Private Sub cmdShowSQL_Click()
frmNotes.Show
frmNotes.txtNote.Text = frmFind.txtQRYHidden

End Sub


Private Sub Form_Load()


Dim Qry As String, i As Integer

' Show the listview Control in Report View
ListView1.View = lvwReport
Qry = "SELECT * FROM " & "[" & TblName & "]"
FillListView Qry, ListView1

For i = 1 To ListView1.ColumnHeaders.Count
cmbFld.List(i - 1) = ListView1.ColumnHeaders(i).Text 'fill the field combo with all the column header names
Next i 'Loop until all fields retrieved
cmbFld.Text = ListView1.ColumnHeaders(1).Text

If ListView1.ListItems.Count > 1 Then
lblRecCount = ListView1.ListItems.Count & _
" Records found in table : " & TblName
Else
lblRecCount = ListView1.ListItems.Count & _
" Record found in table : " & TblName
End If


ErrHnd:

    Exit Sub
 

End Sub

Private Sub Form_Unload(Cancel As Integer)

If oRs.State = 1 Then
oRs.Close
Set oRs = Nothing
End If
If oConn.State = 1 Then oConn.Close



End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHand

Dim x
x = ListView1.ColumnHeaders(ColumnHeader.Index).Text
lblFldLike.Caption = x & " is like : "
lblFldLike.Visible = True: txtSearch.Visible = True
txtSearch.Left = lblFldLike.Left + lblFldLike.Width + 60
txtSearch.Text = ""

Me.Refresh

Dim nIndex As Integer, TagData As Integer

'get the column that was clicked
 nIndex = ColumnHeader.Index - 1
 
 'get the tag value of each column assigned during the load form routine
 TagData = ListView1.ColumnHeaders(ColumnHeader.Index).Tag
  
  With ListView1
    
    Select Case TagData
      Case 2, 3, 4, 5, 6, 17  'Numeric
      
        SortListView nIndex, ListView1, SortNumeric
      
      Case 202, 203, 11  'Text
        SortListView nIndex, ListView1, SortText

      Case 7  'Date
        SortListView nIndex, ListView1, SortDate
    
    Case Else  'if any other values are given, then end
    
    Exit Sub
    
    End Select
  End With
   
Exit Sub

ErrorHand:
MsgBox "Error, Column may have no items"

End Sub

Private Sub txtCriteria1_KeyDown(KeyCode As Integer, Shift As Integer)

If IsNull(txtCriteria1) = False Then
If KeyCode = vbKeyReturn Then cmdRunQuery_Click
End If

End Sub

Private Sub txtCriteria2_KeyDown(KeyCode As Integer, Shift As Integer)

'If IsNull(txtCriteria2) = False Then
If KeyCode = vbKeyReturn Then cmdRunQuery_Click
'End If

End Sub

Private Sub txtSearch_Change()

LvFind ListView1, (Trim$(txtSearch.Text)), txtSearch

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

Execute = True
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
Execute = False
End If

If KeyCode = vbKeyDown Then
ListView1.SetFocus
End If

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)

 KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
 
End Sub


