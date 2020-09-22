Attribute VB_Name = "modLvUtilities"
Option Explicit


Private Declare Function LockWindowUpdate Lib "user32.dll" _
  (ByVal hWnd As Long) As Long

Public Execute As Boolean

Public Enum lvxSortType
  SortText = 0
  SortNumeric = 1
  SortDate = 2
  SortHHMM = 3
  SortHHMMSS = 4
  SortFileDateTime = 5
End Enum

'Declares for FillListView
Public oConn As New ADODB.Connection
Public oRs As New ADODB.Recordset
Dim strConn As String
Public TblName As String, DbPath As String

Public Function FillListView(iQry As String, Optional LView As ListView)

On Error GoTo ErrorHandler

Screen.MousePointer = vbHourglass


Dim MySql As String
'TblName = Dta.RecordSource
'MySql = "SELECT * FROM " & "[" & TblName & "]"

strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DbPath
    
oConn.Open strConn
oRs.CursorLocation = adUseClient
'oRs.Open "Select * from [Customers]", oConn, adOpenStatic, adLockOptimistic
oRs.Open iQry, oConn, adOpenStatic, adLockOptimistic

frmFind.Show
frmFind.Refresh


Dim i As Byte 'Variable used for each Field name and contents

'Fill the Listview headers with the Field names
With LView.ColumnHeaders
   .Clear
   For i = 1 To oRs.Fields.Count
  .Add , , oRs.Fields.Item(i - 1).Name
' Get Data type for each field in Recordset
   Dim Dt As Integer
   Dt = oRs.Fields.Item(i - 1).Type
   LView.ColumnHeaders(i).Tag = Dt 'Now add the Tag data to each column
'   cmbFld.List(i - 1) = LView.ColumnHeaders(i).Text 'fill the field combo with all the column header names
   
   Next i 'Loop until all fields retrieved
'   cmbFld.Text = LView.ColumnHeaders(1).Text
End With 'Stop
 

'Get all the data from the fields in the recordset
Dim itmX As ListItem

        With LView.ListItems
.Clear
        Do Until oRs.eof
                
 Set itmX = .Add(, , oRs(0))
    For i = 2 To oRs.Fields.Count
    If IsNull(oRs(i - 1)) = False Then
        itmX.SubItems(i - 1) = oRs(i - 1)
    Else
    itmX.SubItems(i - 1) = ""
    
End If
Next i
    oRs.MoveNext
        Loop 'Keep going until all data is retrieved
End With

             
'As the Listview loads with the first column sorted, then set
'it's icon to show that.
LView.ColumnHeaders(1).Icon = "Up"
'lblRecCount = LView.ListItems.Count & _
'" Records found in table : " & TblName

Screen.MousePointer = vbDefault


Exit Function

ErrorHandler:

Screen.MousePointer = vbDefault

Dim msg As String
msg = "Please make sure the criteria you entered is appropriate for the" & vbCrLf & _
      "Filed you chose. i.e Numeric Criteria for Numeric Fields and" & vbCrLf & _
      "Text Criteria for Text fields" & vbCrLf & vbCrLf & _
      "Also note: I havn't sorted out Date SQL Yet, Can anyone help here?"
      
Call MsgBox("Runtime error " & Err.Number & ": " & vbCrLf & vbCrLf & _
    msg & vbCrLf & Err.Description, vbOKOnly)
    
    


End Function


Public Sub SortListView(ByVal Index As Integer, _
  ByVal CurrentListView As ListView, _
  Optional vSortType As lvxSortType = SortText)

'On Error GoTo ErrorHandler

  Dim i As Integer
  Dim strFormat As String
  Dim strData() As String
  Dim lRet As Long
  Dim ColHdrName As String
  Dim ColHdrPos As Integer
  'On Error GoTo ErrorHandler
  
'ColHdrName = CurrentListView.ColumnHeaders(Index).Text
'ColHdrPos = CurrentListView.ColumnHeaders(CurrentListView.SortOrder)

  Select Case vSortType

    Case SortText
      With CurrentListView
      
        .Sorted = True

                If .SortKey = Index Then
                    .SortOrder = 1 - .SortOrder
                Else
                    .SortKey = Index
                    .SortOrder = 0
                End If
        

GoSub Clear_Column_Header_Icon:
        

'If .SortOrder = lvwAscending Then
'.ColumnHeaders(Index + 1).Icon = "Up"
'Else
'.ColumnHeaders(Index + 1).Icon = "Down"
'End If

        
        Exit Sub
      
      End With
    
    
    
    Case SortNumeric
      strFormat = String(30, "0") & "." & String(30, "0")
    
    
    Case SortDate
      strFormat = String(2, "0") & "." & String(2, "0") & "." & String(4, "0")
      
    
    Case SortHHMM
      strFormat = "hh:mm"
      
    
    Case SortHHMMSS
      strFormat = "hh:mm:ss"
      
    
    Case SortFileDateTime
      strFormat = String(2, "0") & "." & String(2, "0") & "." & String(4, "0")
      strFormat = strFormat & " " & "hh:mm:ss"
      
    
    Case Else
      Exit Sub
  End Select
    
  
  lRet = LockWindowUpdate(CurrentListView.Parent.hWnd)
  If lRet = 0& Then
    Call MsgBox("Can't lock window " & _
      CurrentListView.Parent.hWnd, _
      vbOKOnly + vbCritical)
    Exit Sub
  End If
    
  With CurrentListView
    With .ListItems
      If (Index > 0) Then
        For i = 1 To .Count
          With .Item(i).ListSubItems(Index)
            .Tag = .Text & vbNullChar & .Tag
            Select Case vSortType
              Case SortNumeric
                If IsNumeric(.Text) Then
                  .Text = Format$(CDbl(.Text), strFormat)
                End If
              Case SortDate
                If IsDate(.Text) Then
                  .Text = Format$(CDate(.Text), strFormat)
                End If
              Case SortHHMM
                .Text = Format$(.Text, strFormat)
              Case SortHHMMSS
                .Text = Format$(.Text, strFormat)
              Case SortFileDateTime
                .Text = Format$(.Text, strFormat)
            End Select
          End With
        Next i
      Else
        For i = 1 To .Count
          With .Item(i)
           
            .Tag = .Text & vbNullChar & .Tag
            Select Case vSortType
              Case SortNumeric
                If IsNumeric(.Text) Then
                  .Text = Format$(CDbl(.Text), strFormat)
                End If
              Case SortDate
                 If IsDate(.Text) Then
                  .Text = Format$(CDate(.Text), strFormat)
                End If
              Case SortHHMM
                .Text = Format$(.Text, strFormat)
              Case SortHHMMSS
                .Text = Format$(.Text, strFormat)
              Case SortFileDateTime
                .Text = Format$(.Text, strFormat)
            End Select
          End With
        Next i
      End If
    End With
        
    
      .Sorted = True


        
                If .SortKey = Index Then
                    .SortOrder = 1 - .SortOrder
                    
                Else
                    .SortKey = Index
                    .SortOrder = 0
                End If
        
GoSub Clear_Column_Header_Icon:
        
    With .ListItems
      If (Index > 0) Then
        For i = 1 To .Count
          With .Item(i).ListSubItems(Index)
            strData = Split(.Tag, vbNullChar)
            .Text = strData(0)
            .Tag = strData(1)
          End With
        Next i
      Else
        For i = 1 To .Count
          With .Item(i)
            strData = Split(.Tag, vbNullChar)
            .Text = strData(0)
            .Tag = strData(1)
          End With
        Next i
      End If
    End With
  End With
        

  lRet = LockWindowUpdate(0&)
  Exit Sub
    
    
    
Clear_Column_Header_Icon:
'This handles the up and down icons in the Column headers

Dim x As Integer
For x = 1 To CurrentListView.ColumnHeaders.Count
If CurrentListView.ColumnHeaders(x).Icon > 0 Then
CurrentListView.ColumnHeaders(x).Icon = 0
CurrentListView.ColumnHeaders(x).Text = CurrentListView.ColumnHeaders(x).Text
End If
Next x

If CurrentListView.SortOrder = lvwAscending Then
CurrentListView.ColumnHeaders(Index + 1).Icon = "Up"
Else
CurrentListView.ColumnHeaders(Index + 1).Icon = "Down"
End If

Return
    
Exit Sub
ErrorHandler:

  Call MsgBox("Runtime error " & Err.Number & ": " & _
    vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Public Sub LvFind(Lvx As ListView, lvxText, txtBox As TextBox)

On Error GoTo ErrorHandler

If Lvx.SortKey < 1 Then GoTo DoFindMain Else GoTo DoFindSub

DoFindMain:
Dim Lvfindtm As ListItem
Dim TempSelStart As Integer
Dim strTemp As String

Set Lvfindtm = Lvx.FindItem(lvxText, lvwText, , lvwPartial)
If Not Lvfindtm Is Nothing Then
Lvfindtm.EnsureVisible
Lvfindtm.Selected = True

If Execute Then
TempSelStart = txtBox.SelStart

If Not txtBox.Text = "" Then

txtBox.SelLength = Len(txtBox.Text) - TempSelStart
    End If
        End If
            End If
Exit Sub

DoFindSub:

Dim LastColumnClicked As Integer
    LastColumnClicked = Lvx.SortKey
'Search Subitems
Dim iSubItemIndex As Integer
Dim i As Integer

iSubItemIndex = Lvx.SortKey
For i = 1 To Lvx.ListItems.Count
If UCase(Lvx.ListItems(i).SubItems(iSubItemIndex)) Like lvxText & "*" Then  'you could also use the LIKE operator
Lvx.ListItems(i).Selected = True
Lvx.ListItems(i).EnsureVisible

Exit For
End If
Next

Exit Sub

ErrorHandler:

  Call MsgBox("Runtime error " & Err.Number & ": " & _
    vbCrLf & Err.Description, vbOKOnly + vbCritical)

End Sub






