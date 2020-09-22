Attribute VB_Name = "modGetData"
Public Function GetDataADOX(strDB As String) As String()

'Type: Public Function
'Name: Get Table Names
'Purpose: Retrieve all table names in an Access database -ADOX Method
'Limitations: VB6 Only
'  Other versions of VB must modify the return value from an array
'Author: Veign
'        http://www.veign.com/information/info_main.html
'Arguments:
'Return Value: A string array of Tables
'Useage: Dim strTables() as String
'        strTables = GetTablesADOX("C:\MyDataBase.mdb")
'Notes: Requires a reference the the MS ADO Ext Library

On Error GoTo Hell

'Temp storage of the Table names
Dim strRetArray() As String
ReDim strRetArray(0) As String

'Create a Catalog object
Dim objCatalog As ADOX.Catalog
Set objCatalog = New ADOX.Catalog

' Open the catalog
objCatalog.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
   "Data Source=" & strDatabase & ";"

' Loop through the tables in the database and print their name
Dim objTable As ADOX.Table
For Each objTable In objCatalog.Tables
   If objTable.Type = "TABLE" Then
        strRetArray(UBound(strRetArray)) = objTable.Name

        'Make room for another item in the array
        ReDim Preserve strRetArray(UBound(strRetArray) + 1)
   End If
Next

'Strip off the last item in the array
ReDim Preserve strRetArray(UBound(strRetArray) - 1)

'Return the array
GetDataADOX = strRetArray

'Dim strTables() As String
'strTables = GetDataADOX("c:\temp\biblio.mdb")
'Debug.Print strTables()

Exit Function

Hell:
    'Error

End Function

