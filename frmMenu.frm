VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMenu 
   Caption         =   "ListView Application"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3540
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4995
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowList 
      Caption         =   "&Show in ListView"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
   End
   Begin MSComctlLib.ListView LvTables 
      Height          =   2895
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   2760
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.mdb"
      DialogTitle     =   "Get Database"
      FileName        =   "*.MDB"
   End
   Begin VB.CommandButton cmdGetDB 
      Caption         =   "Get Database"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblTablesFound 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()

End

End Sub

Private Sub cmdGetDB_Click()

On Error GoTo ErrorHandler

cdl1.InitDir = App.Path
'cdl1.Filter = "*.mdb"
cdl1.ShowOpen
Data2.DatabaseName = cdl1.FileName


LvTables.View = lvwReport
With LvTables.ColumnHeaders
   .Clear
     .Add Text:="Table"
End With

LvTables.ListItems.Clear

Dim strRetArray() As String
ReDim strRetArray(0) As String

'Create a Catalog object
Dim objCatalog As ADOX.Catalog
Set objCatalog = New ADOX.Catalog

' Open the catalog
If cdl1.FileName = "" Then Exit Sub

objCatalog.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
   "Data Source=" & cdl1.FileName

' Loop through the tables in the database and print their name


Dim objTable As ADOX.Table


For Each objTable In objCatalog.Tables
      
      If objTable.Type = "TABLE" Then
        
Dim x As Integer
x = x + 1
'MsgBox x
        
With LvTables.ListItems
.Add (x), , objTable.Name
End With
        
        strRetArray(UBound(strRetArray)) = objTable.Name

        'Make room for another item in the array
        ReDim Preserve strRetArray(UBound(strRetArray) + 1)
   End If
Next



'Strip off the last item in the array
ReDim Preserve strRetArray(UBound(strRetArray) - 1)


LvTables.SetFocus

lblTablesFound.Visible = True
lblTablesFound.Caption = "Tables found in : " & cdl1.FileTitle

Exit Sub

ErrorHandler:

  'Call MsgBox("Runtime error " & Err.Number & ": " & _
   ' vbCrLf & Err.Description, vbOKOnly + vbCritical)
    
End Sub


Private Sub cmdShowList_Click()

On Error GoTo ErrHand

Data2.RecordSource = LvTables.SelectedItem

If LvTables.ListItems.Count = 0 Then
MsgBox "Please Click 'Get Database' and point to a valid .mdb on your Computer"
Exit Sub
End If

TblName = Data2.RecordSource
DbPath = cdl1.FileName

frmFind.Show modal, Me

Exit Sub

ErrHand:
MsgBox "An error has occured, Please get a database first"

Exit Sub


End Sub


Private Sub LvTables_DblClick()

cmdShowList_Click

End Sub
