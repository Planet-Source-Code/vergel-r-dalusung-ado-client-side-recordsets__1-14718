VERSION 5.00
Begin VB.Form frmContacts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Postal Mailing List Maintenance"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8910
   Icon            =   "frmContacts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   5760
      Left            =   255
      TabIndex        =   14
      Top             =   -15
      Width           =   8415
      Begin VB.CommandButton cmdUpdateBatch 
         Caption         =   "Reconnect and Update Database"
         Height          =   435
         Left            =   180
         TabIndex        =   16
         Top             =   5280
         Width           =   1905
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   ">>|"
         Height          =   480
         Index           =   9
         Left            =   5865
         TabIndex        =   17
         Top             =   4170
         Width           =   1005
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   ">"
         Height          =   480
         Index           =   8
         Left            =   4845
         TabIndex        =   15
         Top             =   4170
         Width           =   1005
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "<"
         Height          =   480
         Index           =   7
         Left            =   3840
         TabIndex        =   13
         Top             =   4170
         Width           =   1005
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "|<<"
         Height          =   480
         Index           =   6
         Left            =   2835
         TabIndex        =   12
         Top             =   4170
         Width           =   1005
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Exit"
         Height          =   480
         Index           =   5
         Left            =   6900
         TabIndex        =   11
         Top             =   3540
         Width           =   1005
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Add New"
         Height          =   480
         Index           =   4
         Left            =   5886
         TabIndex        =   10
         Top             =   3540
         Width           =   1005
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Save"
         Height          =   480
         Index           =   3
         Left            =   4872
         TabIndex        =   9
         ToolTipText     =   "Changes in a disconnected recordset are not saved to file until .UpdateBatch"
         Top             =   3540
         Width           =   1005
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Edit"
         Height          =   480
         Index           =   2
         Left            =   3858
         TabIndex        =   8
         Top             =   3540
         Width           =   1005
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Cancel"
         Height          =   480
         Index           =   1
         Left            =   2844
         TabIndex        =   7
         Top             =   3540
         Width           =   1005
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Delete"
         Height          =   480
         Index           =   0
         Left            =   1830
         TabIndex        =   6
         Top             =   3540
         Width           =   1005
      End
      Begin VB.ComboBox cboContactName 
         Height          =   315
         Left            =   1845
         TabIndex        =   0
         Text            =   "(enter contact person's name)"
         Top             =   915
         Width           =   6135
      End
      Begin VB.TextBox txtState 
         Height          =   315
         Left            =   1830
         TabIndex        =   4
         Text            =   "(enter State)"
         Top             =   2820
         Width           =   2175
      End
      Begin VB.TextBox txtAddressLine1 
         Height          =   315
         Left            =   1830
         TabIndex        =   2
         Text            =   "(number and street)"
         Top             =   1860
         Width           =   6135
      End
      Begin VB.TextBox txtCity 
         Height          =   315
         Left            =   1830
         TabIndex        =   3
         Text            =   "(enter city name)"
         Top             =   2325
         Width           =   6135
      End
      Begin VB.TextBox txtZipCode 
         Height          =   315
         Left            =   5790
         TabIndex        =   5
         Text            =   "(enter zip code)"
         Top             =   2820
         Width           =   2175
      End
      Begin VB.TextBox txtCompany 
         Height          =   315
         Left            =   1830
         TabIndex        =   1
         Text            =   "(enter your dba, if any)"
         Top             =   1380
         Width           =   6135
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mailing List Maintenance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   285
         TabIndex        =   26
         Top             =   225
         Width           =   3150
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<-- Click Button to Reconnect and do Batch Update on the Disconnected Recordset."
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2175
         TabIndex        =   25
         Top             =   5295
         Width           =   6195
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status: Recordset is now Disconnected from .mdb file to free up open files and network connection if applicable. "
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         TabIndex        =   24
         Top             =   4890
         Width           =   8265
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip Code:"
         Height          =   255
         Left            =   4545
         TabIndex        =   23
         Top             =   2865
         Width           =   975
      End
      Begin VB.Label lblContactName 
         Caption         =   "Contact Name:"
         Height          =   255
         Left            =   375
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblStreetAddress 
         Caption         =   "Street Address:"
         Height          =   255
         Left            =   375
         TabIndex        =   21
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblCity 
         Caption         =   "City:"
         Height          =   255
         Left            =   375
         TabIndex        =   20
         Top             =   2385
         Width           =   1335
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company Name:"
         Height          =   255
         Left            =   375
         TabIndex        =   19
         Top             =   1410
         Width           =   1335
      End
      Begin VB.Label lblState 
         Caption         =   "State:"
         Height          =   255
         Left            =   375
         TabIndex        =   18
         Top             =   2865
         Width           =   1335
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Included as Project References
'Microsoft Scripting Runtime
'Microsoft ADO Ext. for DDL and Security
'Microsoft ActiveX Data Object 2.x Library
'Microsoft DataBinding collection

'Original codes by Legrev3@aol.com
'ADOX database creation, ADO databinding, Disconnected Recordsets demo
'Submitted for downloading Jan. 24, 2001

Dim cnContacts As ADODB.Connection
Dim rsContacts As ADODB.Recordset
Dim OneRec As UserDefRec

Dim blnAdd As Boolean           'set to true if adding
Dim blnEdit As Boolean          'set to true if editing
Dim blnUpdated As Boolean       'set to true if modifcations are written to database
Dim blnExiting As Boolean       'set to true if closing app

Dim strSearch As String
Dim strEditRec As String
Dim vntBookMark As Variant
Dim intResponse As Integer

Private Enum cmdButtons
    DeleteButton = 0
    CancelButton = 1
    EditButton = 2
    SaveButton = 3
    AddNewButton = 4
    ExitButton = 5
    MoveFirstButton = 6
    MovePreviousButton = 7
    MoveNextButton = 8
    MoveLastButton = 9
End Enum

Private Sub Form_Load()
    Dim blnFileExists As Boolean
    Dim strDbName As String
    
    strDbName = "Contacts.mdb"                      'database name to use
    
    'strFilespec and strConn are declared in modCreateDatabase.bas
    'in a client-server environment, the db path would be on the network server
    'using an ODBC compliant provider.
    'this demo was modified to run on a stand-alone computer.
    
    strFilespec = App.Path & "\Data\" & strDbName
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilespec & ";"
    
    blnFileExists = DatabaseExist()        'fn to see if Contacts.mdb exist
    
    If blnFileExists = False Then
        intResponse = MsgBox(strFilespec & " does not exist and will now be created.", vbInformation + vbOKCancel)
        If intResponse = vbCancel Then Unload Me
        Call CreateDataFile                'if not exist create .mdb
        Call CreateMailListTable                    'create table MailList
    End If
    
    Set cnContacts = New ADODB.Connection       'use ADO connection
    Set rsContacts = New ADODB.Recordset        'and ADO recordset
    
    If cnContacts.State = adStateClosed Then    'newly-created db may be open so
        cnContacts.CursorLocation = adUseClient 'create client side
        cnContacts.Open strConn                 'check first before opening
    End If
    
    rsContacts.Open "SELECT * FROM MailList ORDER BY ContactName", cnContacts, adOpenStatic, adLockBatchOptimistic, adCmdText
        
    'populate combo box with names
    Do Until rsContacts.EOF
        If Not IsNull(rsContacts!ContactName) Then cboContactName.AddItem rsContacts!ContactName
        rsContacts.MoveNext
    Loop
    
    'data binding - do not bind cboContactName to enable search ability
    Set txtCompany.DataSource = rsContacts
    txtCompany.DataField = "CompanyName"
    Set txtAddressLine1.DataSource = rsContacts
    txtAddressLine1.DataField = "AddressLine1"
    Set txtCity.DataSource = rsContacts
    txtCity.DataField = "City"
    Set txtState.DataSource = rsContacts
    txtState.DataField = "State"
    Set txtZipCode.DataSource = rsContacts
    txtZipCode.DataField = "ZipCode"
    
    If cboContactName.ListCount > 0 Then
        rsContacts.MoveFirst
        cboContactName.Text = rsContacts!ContactName
    End If
    
    DoEvents
    'Create a disconnected recordset
    'after getting the records, the client now disconnects.
    'this is the essence of disconnected recordsets
    
    Set rsContacts.ActiveConnection = Nothing
    cnContacts.Close
    
    blnUpdated = True       'flag the fact that batch update is not yet necessary
End Sub


Private Sub cmdButton_Click(Index As Integer)
'determine which button is clicked and do appropriate action
    Select Case Index
        Case DeleteButton
            Dim strTemp As String
            Dim i As Integer
            
            If cboContactName.ListCount = 1 Then
                MsgBox "Deleting ALL records will cause an Update error." & vbCr & _
                        "Sorry, last record will not be deleted in this simulation.", vbCritical + vbOKOnly
                Exit Sub
            End If
            
            If rsContacts.EOF Or rsContacts.BOF Then
                MsgBox "There is no record to delete.", vbExclamation + vbOKOnly
                Exit Sub
            End If
            
            intResponse = MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo)
            If intResponse = vbNo Then Exit Sub
            
            strTemp = Trim$(cboContactName.Text)
            
            On Error GoTo DeleteError:
            rsContacts.Delete
            On Error GoTo 0
                        
            For i = 0 To cboContactName.ListCount
                If cboContactName.List(i) = strTemp Then
                    cboContactName.RemoveItem i
                    Exit For
                End If
            Next i
                            
            rsContacts.MoveNext
            If rsContacts.EOF Then rsContacts.MoveFirst
            
            blnUpdated = False   'we have something to write to db from disconnected recordset
            cboContactName.Text = rsContacts!ContactName
            cboContactName.SetFocus
            DoEvents
                        
        Case CancelButton
            If blnEdit = False And blnAdd = False Then
                MsgBox "There is no Edit or Add activity to Cancel.", vbExclamation
                Call EnableButtons("YYYYYYYYYY")
                Exit Sub
            End If
            
            Call EnableButtons("YYYYYYYYYY")
            
            If blnEdit = True Then
                Call ClearText
                Call RestoreCurrent
                cboContactName = rsContacts!ContactName
                blnEdit = False
                cboContactName.SetFocus
                Exit Sub
            End If
            
            
            On Error Resume Next
            
            rsContacts.CancelUpdate
            If vntBookMark <> "" Then rsContacts.Bookmark = vntBookMark
            vntBookMark = ""
            cboContactName.Text = rsContacts!ContactName
            cboContactName.SetFocus
            blnAdd = False
            On Error GoTo 0
            DoEvents
                        
        Case EditButton
            If cboContactName.ListCount = 0 Then
                MsgBox "Contacts tables is empty."
                cboContactName.SetFocus
                Exit Sub
            End If
            
            blnEdit = True                  'set blnEdit flag - we are editing
            
            Call SaveCurrent                'in case of cancel
            
            strEditRec = cboContactName
            Call EnableButtons("NYNYNNNNNN")        'set applicable buttons to Y
            cboContactName.Text = rsContacts!ContactName
            txtCompany.SetFocus
            
        Case SaveButton
            If blnAdd = False And blnEdit = False Then
                MsgBox "There is no edited or added record to save.", vbExclamation
                Exit Sub
            End If
            
            Call EnableButtons("YYYYYYYYYY")        'all buttons enabled
            
            If blnEdit = True Then                  'Edit was clicked so
                blnEdit = False                     'reset blnEdit flag
                If strEditRec <> cboContactName Then
                    MsgBox "ContactName is a primary field and may not be edited.", vbExclamation + vbOKOnly
                    Call RestoreCurrent
                    cboContactName = rsContacts!ContactName
                    Exit Sub
                End If
            Else
                rsContacts!ContactName = cboContactName.Text
                rsContacts.Update
                rsContacts.MoveLast                         'added new record
                cboContactName.AddItem cboContactName.Text  'update combo box
                blnAdd = False                              'reset blnAdd flag
            End If
            On Error GoTo 0
            
            blnUpdated = False   'we have something to write to db from disconnected recordset
            
        Case AddNewButton
            blnAdd = True                       'set blnAdd flag - we are adding
            vntBookMark = ""
            On Error Resume Next
            vntBookMark = rsContacts.Bookmark   'in case Cancel is clicked
            rsContacts.AddNew
            Call EnableButtons("NYNYNNNNNN")
            Call ClearText
            cboContactName.SetFocus
            On Error GoTo 0
            
        Case ExitButton
            Unload Me
            
'Navigation buttons
        Case MoveFirstButton
            If cboContactName.ListCount = 0 Then
                MsgBox "Contacts tables is empty.", vbExclamation
            Else
                rsContacts.MoveFirst
                cboContactName.Text = rsContacts!ContactName
            End If
            
        Case MovePreviousButton
            If cboContactName.ListCount = 0 Then
                MsgBox "Contacts tables is empty.", vbExclamation
            Else
                rsContacts.MovePrevious
                If rsContacts.BOF Then
                    rsContacts.MoveFirst
                    MsgBox "The current record is the first on file."
                End If
                cboContactName.Text = rsContacts!ContactName
            End If
            
        Case MoveNextButton
            If cboContactName.ListCount = 0 Then
                MsgBox "Contacts tables is empty.", vbExclamation
            Else
                rsContacts.MoveNext
                If rsContacts.EOF Then
                    rsContacts.MoveLast
                    MsgBox "The current record is the last on file."
                End If
                cboContactName.Text = rsContacts!ContactName
            End If
            
        Case MoveLastButton
            If cboContactName.ListCount = 0 Then
                MsgBox "Contacts tables is empty.", vbExclamation
            Else
                rsContacts.MoveLast
                cboContactName.Text = rsContacts!ContactName
            End If
    End Select
Exit Sub

DeleteError:
    MsgBox "There has been a delete error. If record is not deleted try again later."
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If blnEdit = False And blnAdd = False Then
        MsgBox "Click Edit to edit records or AddNew to enter new records.", vbOKOnly + vbInformation
        KeyAscii = 0
    End If
End Sub

Public Sub EnableButtons(strYN As String)
'parameter Y to enable no to disable cmdButton(0) to (10)
'position of Y or N on strYN corresponds to button index
    Dim intIndex As Integer
    Dim intAllButtons As Integer
    
    strYN = Trim$(strYN)
    intAllButtons = Len(strYN)
    
    For intIndex = 1 To intAllButtons
        cmdButton(intIndex - 1).Enabled = True      'default is enabled
        If (Mid$(strYN, intIndex, 1) = "N") Then cmdButton(intIndex - 1).Enabled = False
    Next intIndex
    
End Sub


Public Sub ClearText()
'clears text of all bound controls - textboxes and combobox
    Dim i As Integer
    
    For i = 1 To Me.Controls.Count - 1
        If (TypeOf Me.Controls(i) Is TextBox) Then
            Me.Controls(i).Text = ""
        ElseIf (TypeOf Me.Controls(i) Is ComboBox) Then
            Me.Controls(i).Text = ""
        End If
    Next i
    cboContactName.SetFocus
End Sub

Private Sub cboContactName_Click()
'fires when a combobox item is clicked or Enter is hit on the combobox
'determines if contact name already exist on file
'if not, then allow to add
    If blnAdd = False Then                  'AddNew was not clicked
        Dim blnFound As Boolean
        strSearch = CStr(cboContactName.Text)
        blnFound = FindRec()
        
        If blnFound = True Then             'contact name is on file, displayed
            cboContactName.SetFocus
            Exit Sub
        End If
    End If
    
'user clicked AddNew or ContactName is not on file
'by setting focus on txtCompany, we are triggering cboContactName_LostFocus()
    txtCompany.SetFocus
    
End Sub

Private Sub cboContactName_KeyPress(KeyAscii As Integer)
'triggers cboContactName.LostFocus()
    If KeyAscii <> Asc(vbCr) Then Exit Sub
    txtCompany.SetFocus
End Sub

Private Sub cboContactName_LostFocus()
'user leaves combobox, determine what is on it except for Cancel button
    If Me.ActiveControl = cmdButton(CancelButton) Then Exit Sub
    Dim intResponse As Integer
    intResponse = Validate
    If intResponse = 1 Then
        cmdButton_Click (CancelButton)      'Cancel
        Exit Sub
    ElseIf intResponse = 2 Then
        cboContactName.SetFocus
        Exit Sub
    End If
    
    ' intResponse = 0 allow to proceed
    If blnAdd = True Then
        Dim i As Integer
        Dim blnFound As Boolean
        blnFound = False
        strSearch = CStr(cboContactName.Text)
        For i = 0 To cboContactName.ListCount
            If cboContactName.List(i) = cboContactName.Text Then
                rsContacts.CancelUpdate
                rsContacts.MoveFirst
                blnAdd = False
                blnFound = FindRec()
                Exit For
            End If
        Next i
        If blnFound = True Then
            cboContactName.Text = rsContacts!ContactName
            Call EnableButtons("YYYYYYYYYY")
            MsgBox "Record Exists.", vbInformation + vbOKOnly
            DoEvents
            cboContactName.SetFocus
            Exit Sub
        End If
    End If
End Sub

Public Sub SaveCurrent()
'save current record if edit is clicked
    With OneRec
        On Error Resume Next
        .strContactName = CStr(rsContacts!ContactName) & ""
        .strCompany = CStr(rsContacts!CompanyName) & ""
        .strAddressLine1 = CStr(rsContacts!AddressLine1) & ""
        .strCity = CStr(rsContacts!City) & ""
        .strState = CStr(rsContacts!State) & ""
        .strZipCode = CStr(rsContacts!ZipCode) & ""
    End With
End Sub

Public Sub RestoreCurrent()
'restore current record if Cancel is clicked after edit
    With rsContacts
        On Error Resume Next
        !ContactName = OneRec.strContactName
        !CompanyName = OneRec.strCompany
        !AddressLine1 = OneRec.strAddressLine1
        !City = OneRec.strCity
        !State = OneRec.strState
        !ZipCode = OneRec.strZipCode
    End With
End Sub

Private Sub cmdUpdateBatch_Click()
    If blnUpdated = True Then
        MsgBox "There are no modifications to save to database file.", vbInformation
        cboContactName.SetFocus
        Exit Sub
    End If

    cnContacts.Open                              'reopen connection
    Set rsContacts.ActiveConnection = cnContacts 'reconnect recordset
    rsContacts.UpdateBatch
    
    If blnExiting = True Then                     'fired from QueryUnload
        Set rsContacts.ActiveConnection = Nothing
        cnContacts.Close
        Exit Sub
    End If
    
    
    Label1 = "Status: Connection reopened and database updated. Closed after update."
    MsgBox "Modifications to Disconnected Recordset has been saved." & vbCr & "Press OK to Close and recreate Disconnected Recordset."
    
    
    'recreate recordset
    rsContacts.Requery
    
    'populate combo box with names
    cboContactName.Clear
    Do Until rsContacts.EOF
        If Not IsNull(rsContacts!ContactName) Then cboContactName.AddItem rsContacts!ContactName
        rsContacts.MoveNext
    Loop
    
    'close connection to create a disconnected recordset
    Set rsContacts.ActiveConnection = Nothing
    cnContacts.Close
    
    'set updated flag
    blnUpdated = True
    blnAdd = False
    blnEdit = False
    rsContacts.MoveFirst
    cboContactName.Text = rsContacts!ContactName
End Sub

Private Function Validate() As Integer
    Dim x As Integer
    Dim strMsg As String
    
    If blnEdit = True Then
        If strEditRec <> cboContactName.Text Then
            strMsg = "The Contact Name field is a primary key and may not be edited."
            GoTo CannotEditName:
        Else
            Validate = 0
            Exit Function
        End If
    End If
    
    'new record is being added
    cboContactName.Text = Trim$(cboContactName.Text)
    If Len(cboContactName.Text) = 0 Then
        strMsg = "The ContactName field is the primary key and is required."
        GoTo RequiredField:
    End If
        
    Validate = 0
    Exit Function

CannotEditName:
'usual for the primary key is a CompanyId or uniqe identifier
'we use the contact name to somewhat shorten code
    MsgBox strMsg, vbExclamation + vbOKOnly
    Validate = 1
    Exit Function
    
RequiredField:
    x = MsgBox(strMsg, vbRetryCancel)
    If x = vbCancel Then
        Validate = 1
    Else
        Validate = 2
    End If
End Function

Public Function FindRec() As Boolean
    Dim strTemp
    
    strTemp = "'" & strSearch & "'"
    On Error Resume Next
    rsContacts.MoveFirst
    
    On Error GoTo ErrorNotOnFile:
    rsContacts.Find "ContactName = " & strTemp, 0, adSearchForward
    
    If rsContacts!ContactName = strSearch Then FindRec = True       'found
    On Error GoTo 0
    Err.Clear
    Exit Function
    
ErrorNotOnFile:
    MsgBox "Error =   " & Err.Number & Err.Description
    FindRec = False      'not found
    DoEvents
    On Error GoTo 0
    Err.Clear
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If blnUpdated = False Then
        intResponse = MsgBox("Do you wish to send your modifications to the database file?", vbQuestion + vbYesNoCancel)
        If intResponse = vbYes Then
            blnExiting = True
            Call cmdUpdateBatch_Click
        ElseIf intResponse = vbNo Then
            Exit Sub
        ElseIf intResponse = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub


'these KeyPress Subs merely allow the enter key to behave like the tab key
'and are not essential in this demo
'the GotFocus codes merely highlight text and again are not essential

Private Sub txtAddressLine1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> Asc(vbCr) Then Exit Sub
    txtAddressLine1 = Trim$(txtAddressLine1)
    txtCity.SetFocus
End Sub

Private Sub txtAddressLine1_GotFocus()
    txtAddressLine1.SelStart = 0
    txtAddressLine1.SelLength = Len(txtAddressLine1)
End Sub

Private Sub txtCity_GotFocus()
    txtCity.SelStart = 0
    txtCity.SelLength = Len(txtCity)
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
    If KeyAscii <> Asc(vbCr) Then Exit Sub
    txtCity = Trim$(txtCity)
    txtState.SetFocus
End Sub

Private Sub txtCompany_GotFocus()
    txtCompany.SelStart = 0
    txtCompany.SelLength = Len(txtCompany)
End Sub

Private Sub txtCompany_KeyPress(KeyAscii As Integer)
    If KeyAscii <> Asc(vbCr) Then Exit Sub
    txtCompany = Trim$(txtCompany)
    txtAddressLine1.SetFocus
End Sub

Private Sub txtState_GotFocus()
    txtState.SelStart = 0
    txtState.SelLength = Len(txtState)
End Sub

Private Sub txtState_KeyPress(KeyAscii As Integer)
    If KeyAscii <> Asc(vbCr) Then Exit Sub
    txtState = Trim$(txtState)
    txtZipCode.SetFocus
End Sub

Private Sub txtZipCode_GotFocus()
    txtZipCode.SelStart = 0
    txtZipCode.SelLength = Len(txtZipCode)
End Sub

Private Sub txtZipCode_KeyPress(KeyAscii As Integer)
    If KeyAscii <> Asc(vbCr) Then Exit Sub
    txtZipCode = Trim$(txtZipCode)
    cmdButton(SaveButton).SetFocus

End Sub
