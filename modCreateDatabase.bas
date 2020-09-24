Attribute VB_Name = "modCreateDatabase"
Option Explicit
'Included as Project References
'Microsoft Scripting Runtime
'Microsoft ADO Ext. for DDL and Security
'Microsoft ActiveX Data Object 2.x Library
'Microsoft DataBinding collection

'Original codes by Legrev3@aol.com
'ADOX database creation, ADO databinding, Disconnected Recordsets demo
'Submitted for downloading Jan. 24, 2001

Public Type UserDefRec
    strContactName As String
    strCompany As String
    strAddressLine1 As String
    strCity As String
    strState As String
    strZipCode As String
End Type

Public Contacts As ADOX.Catalog
Public strFilespec As String
Public strConn As String

'If not exist create Access2000 database file Contacts.mdb
Public Sub CreateDataFile()
    Set Contacts = New ADOX.Catalog
    Contacts.Create strConn
End Sub

'Create MailList table
Public Sub CreateMailListTable()

    Dim tbl As ADOX.Table
    Dim key As ADOX.key
        
    Set tbl = New ADOX.Table
'TableName.Columns.Append "ColumnName" is an ADOX
'method of creating columns of a table
    With tbl
        Set .ParentCatalog = Contacts
        .Name = "MailList"
        .Columns.Append "ContactName", adVarWChar, 70
        .Columns.Append "CompanyName", adVarWChar, 70
        .Columns("CompanyName").Attributes = adColNullable
        .Columns.Append "AddressLine1", adVarWChar, 70
        .Columns("AddressLine1").Attributes = adColNullable
        .Columns.Append "City", adVarWChar, 70
        .Columns("City").Attributes = adColNullable
        .Columns.Append "State", adVarWChar, 30
        .Columns("State").Attributes = adColNullable
        .Columns.Append "ZipCode", adVarWChar, 30
        .Columns("ZipCode").Attributes = adColNullable
'for demo purposes use ContactName as key. ContactId is usual
        .Keys.Append "PrimaryKey", adKeyPrimary, "ContactName"
    End With
'add table to database .mdb file
    Contacts.Tables.Append tbl
End Sub

'determine database file exists in App.Path & "\Data\"
'returns True if database exists
Public Function DatabaseExist() As Boolean
    Dim fso As FileSystemObject
    Dim fldr As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
'check if folder \Data exists, if not, then create
    fldr = App.Path & "\Data"
    If Not (fso.FolderExists(fldr)) Then
        fso.CreateFolder (fldr)
    End If
    
    DatabaseExist = fso.FileExists(strFilespec)
End Function

