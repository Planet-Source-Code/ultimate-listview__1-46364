Attribute VB_Name = "ExportUtils"
Public Function ExportCSV(myFileTitle As String)

On Error GoTo ErrHand

   Dim sData As String
   oRs.MoveFirst
   sData = oRs.GetString(adClipString, , ",", vbCr, vbNullString)
   Open myFileTitle For Output As #1
   Print #1, sData
   Close #1
    
Exit Function

ErrHand:
Err.Clear
'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
frmMenu.Show

End Function


Public Function ExportToExcel(ByRef pListview As MSComctlLib.ListView, _
    ByVal pFilename As String)
    Dim CN As Object
    Dim CAT As Object
    Dim TBL As Object
    Dim COL As Object
    Dim strConnection As String
    Dim AListItem As MSComctlLib.ListItem
    Dim AColumnHeader As MSComctlLib.ColumnHeader
    Dim RS As Object
    Dim intLoop As Integer
    Dim intLoop2 As Integer
    
    On Error GoTo ErrHandler
    
    ' Make sure everything is ok with the inputs before
    ' continuing.
    ' pListView
    If pListview.View <> lvwReport Then
        MsgBox "Listview must be in Report mode.", _
            vbCritical + vbOKOnly, "ExportListview"
        GoTo NotSuccessful
    End If
    ' pFilename
    If Trim$(pFilename) = vbNullString Then
        MsgBox "No filename given.", vbCritical + vbOKOnly, _
            "ExportListview"
        GoTo NotSuccessful
    End If
    ' **********
    Set CN = CreateObject("ADODB.Connection")
    
    ' Create a connection to the Excel file using Jet's ISAM
    strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Extended Properties=Excel 8.0;" & _
        "Data Source=" & pFilename
    CN.Open strConnection
    
    ' No need to create a workbook if the spreadsheet already exists
    If Append Then GoTo AlreadyExists
    
    ' Create a Excel Workbook and set the connection to CN
    Set CAT = CreateObject("ADOX.Catalog")
    CAT.ActiveConnection = CN
    
    ' Create a worksheet for the cat
    Set TBL = CreateObject("ADOX.Table")
    
'    If TBL.Name = "Sheet1" Then
'    MsgBox "hi"
'    CAT.Tables.Delete TBL.Name
'    End If
    
    
    TBL.Name = "Sheet1"
    
    
    ' Do the column headers
    
    For Each AColumnHeader In pListview.ColumnHeaders
        Set COL = CreateObject("ADOX.Column")
        COL.Type = 130  ' adWChar
        COL.Name = AColumnHeader.Text
        TBL.Columns.Append COL
        Set COL = Nothing
    Next AColumnHeader
    
   ' Add this worksheet to the workbook

    If CAT.Tables.Count = 2 Then
    CAT.Tables.Delete TBL.Name
    End If

    CAT.Tables.Append TBL


AlreadyExists:
    Set RS = CreateObject("ADODB.Recordset")
    
    ' open the excel file that was just created as a recordset
    ' so we can add records.
    RS.Open TBL.Name, CN, 1, 3
    
    ' Grab every listitem out of the listview control
    
    For Each AListItem In pListview.ListItems
        ' Listitem and then all subitems
        RS.AddNew
        RS.Fields(0) = AListItem.Text
        ' subitems
        For intLoop = 1 To RS.Fields.Count - 1
            RS.Fields(intLoop) = AListItem.SubItems(intLoop)
        Next intLoop
        RS.Update
    Next AListItem
    
        

    ' Mark as success
    ExportListview = True
    GoTo CloseAndNothing
    
NotSuccessful:
    ExportListview = False
    
    ' clear all objects and exit
CloseAndNothing:
    On Error Resume Next
    RS.Close
    CN.Close
    Set CAT = Nothing
    Set CN = Nothing
    Set TBL = Nothing
    Set COL = Nothing
    Set AListItem = Nothing
    Set AColumnHeader = Nothing
    Set RS = Nothing
    Exit Function
    
ErrHandler:
    ' simply raise the error to the client
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    GoTo CloseAndNothing
End Function


'Clipboard
Public Sub SendToClipboard(ByVal ListViewObj _
As MSComctlLib.ListView)
Dim ListItemObj As MSComctlLib.ListItem
Dim ListSubItemObj As MSComctlLib.ListSubItem
Dim ColumnHeaderObj As _
MSComctlLib.ColumnHeader

If ListViewObj.ListItems.Count > 1000 Then
Dim MyMsg
MyMsg = MsgBox("More than 1000 records, Are you sure?", vbQuestion + _
    vbYesNo + vbDefaultButton1, "Warning")

If MyMsg = vbNo Then Exit Sub
End If

Dim ClipboardText As String
Dim ClipboardLine As String
Clipboard.Clear
For Each ColumnHeaderObj In _
ListViewObj.ColumnHeaders
Select Case ColumnHeaderObj.Index
Case 1
ClipboardText = ColumnHeaderObj.Text
Case Else
ClipboardText = ClipboardText & _
vbTab & ColumnHeaderObj.Text
End Select
Next ColumnHeaderObj
For Each ListItemObj In _
ListViewObj.ListItems
ClipboardLine = ListItemObj.Text
For Each ListSubItemObj In ListItemObj.ListSubItems
ClipboardLine = ClipboardLine & _
vbTab & ListSubItemObj.Text
Next ListSubItemObj
ClipboardText = ClipboardText & vbCrLf _
& ClipboardLine
Next ListItemObj
Clipboard.SetText ClipboardText

MsgBox "Data has been copied to the Windows Clipboard"

End Sub

