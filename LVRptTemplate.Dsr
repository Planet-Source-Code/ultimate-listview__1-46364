VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} LVTemplate 
   Caption         =   "Listview Report"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "LVRptTemplate.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "LVRptTemplate.dsx":5C12
End
Attribute VB_Name = "LVTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LView As ListView
Private bDone As Boolean
Dim ColHdr As ColumnHeader


Property Set Grid(Lv1 As ListView)

Dim ctl As Object
Dim iLeft As Integer
Dim i As Integer
    
Set LView = Lv1

'Setup The Header Section
For i = 1 To LView.ColumnHeaders.Count
    Set ctl = PageHeader.Controls.Add("DDActiveReports2.Field")
    Fields.Add ctl.Name
    ctl.DataField = ctl.Name
'If the user has rearranged the column headers on the Listview
'this will reflect that on the report (Header Part)
        For Each ColHdr In LView.ColumnHeaders
            If ColHdr.Position = i Then 'check the position of the column header
                Fields(ctl.Name).Value = LView.ColumnHeaders(ColHdr.Index).Text
                ctl.Font.Name = "Verdana"
                ctl.Font.Size = 8
                ctl.Font.Bold = True
                ctl.Left = iLeft
                ctl.Top = 0
                ctl.Width = LView.ColumnHeaders.Item(LView.ColumnHeaders(ColHdr.Index).Index).Width
                ctl.Tag = i
            End If
        Next ColHdr
    iLeft = iLeft + ctl.Width + 144
    PrintWidth = iLeft
Next i

i = 0: iLeft = 0 'reset our variables

'Setup the Body Fields, again checking the column orde rincase the user has rearranged them
For i = 1 To LView.ColumnHeaders.Count
    Set ctl = Detail.Controls.Add("DDActiveReports2.Field")
        For Each ColHdr In LView.ColumnHeaders
                If ColHdr.Position = i Then
                    ctl.Name = "F" & LView.ColumnHeaders(ColHdr.Index).Index 'name the fields F1, F2 F3 etc
                    Fields.Add "F" & LView.ColumnHeaders(ColHdr.Index).Index 'add the fields
                    ctl.DataField = "F" & LView.ColumnHeaders(ColHdr.Index).Index 'again
                    ctl.Tag = LView.ColumnHeaders(ColHdr.Index).Index 'set the tag here as the the col hdr position,
                                                                  'this is used in getting the data in the function below in the right order.
                    ctl.Left = iLeft
                    ctl.Top = 0
                    ctl.Width = LView.ColumnHeaders.Item(LView.ColumnHeaders(ColHdr.Index).Index).Width
                End If
        Next ColHdr
    iLeft = iLeft + ctl.Width + 144
    PrintWidth = iLeft
Next i
        
End Property

Private Sub ActiveReport_FetchData(eof As Boolean)


Static iRow As Integer
Dim ctl As Object
Dim i As Integer, x As Integer

'Screen.MousePointer = vbHourglass
'Fill the Data
If iRow < LView.ListItems.Count Then 'get the count of all the items in the listview
    For Each ctl In Detail.Controls
        If ctl.Tag = 1 Then
        Fields(ctl.Name).Value = LView.ListItems.Item(iRow + 1).Text
        Else
        Fields(ctl.Name).Value = LView.ListItems.Item(iRow + 1).ListSubItems(ctl.Tag - 1).Text
        End If

Next ctl
iRow = iRow + 1
eof = False
End If

'Screen.MousePointer = vbDefault
End Sub


