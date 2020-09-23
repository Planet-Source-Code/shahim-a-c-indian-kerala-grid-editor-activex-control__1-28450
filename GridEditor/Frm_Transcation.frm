VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmConnection 
   Caption         =   "Connection Parameters"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   8940
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "Show All Types"
      Height          =   240
      Left            =   5385
      TabIndex        =   6
      Top             =   525
      Width           =   1515
   End
   Begin VB.OptionButton OptEditor 
      Caption         =   "Use Grideditor"
      Height          =   210
      Left            =   2625
      TabIndex        =   5
      Top             =   540
      Width           =   1380
   End
   Begin VB.OptionButton OptPick 
      Caption         =   "Use PickList"
      Height          =   210
      Left            =   4005
      TabIndex        =   4
      Top             =   540
      Value           =   -1  'True
      Width           =   1380
   End
   Begin MSComctlLib.ListView LstTable 
      Height          =   5325
      Left            =   30
      TabIndex        =   3
      Top             =   840
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   9393
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Table Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date Created"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date Modified"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox TxtConnectionString 
      Height          =   315
      Left            =   2625
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   135
      Width           =   5685
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Transcation.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Transcation.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Transcation.frx":02BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Transcation.frx":070E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Connection String :"
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   165
      Width           =   1350
   End
End
Attribute VB_Name = "FrmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
On Error GoTo errH
If Conn.State = 0 Then Exit Sub
Dim Catlg As New ADOX.Catalog
Catlg.ActiveConnection = Conn
Dim tbl As Table
Dim Lst As ListItem

LstTable.ListItems.Clear
If Check1.Value = vbUnchecked Then
    For Each tbl In Catlg.Tables
    '---Get only tables
        If tbl.Type = "TABLE" Then
            Set Lst = LstTable.ListItems.Add(, , tbl.Name, 1, 1)
            Lst.SubItems(1) = IIf(IsNull(tbl.DateCreated), "", tbl.DateCreated) 'Lst.SubItems(1) = tbl.DateCreated
            Lst.SubItems(2) = IIf(IsNull(tbl.DateModified), "", tbl.DateModified)
            Lst.SubItems(3) = tbl.Type
        End If
    Next
Else
    For Each tbl In Catlg.Tables
    '---Get only tables
        Set Lst = LstTable.ListItems.Add(, , tbl.Name, 1, 1)
        Lst.SubItems(1) = IIf(IsNull(tbl.DateCreated), "", tbl.DateCreated) 'Lst.SubItems(1) = tbl.DateCreated
        Lst.SubItems(2) = IIf(IsNull(tbl.DateModified), "", tbl.DateModified)
        Lst.SubItems(3) = tbl.Type
    Next
End If
LstTable.ColumnHeaders(1).Width = 4000
LstTable.ColumnHeaders(2).Width = 3000
LstTable.ColumnHeaders(3).Width = 3000
'---Sort accordng to type
LstTable.Sorted = True
LstTable.SortKey = 3
LstTable.SortOrder = lvwAscending
LstTable.Sorted = False
Exit Sub
errH:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub CmdBrowse_Click()
On Error GoTo errH
Dim Obj As Object               '---Reference to the connectionstring
Dim Dlink As New DataLinks      '---Reference to datalink
Dim Catlg As New Catalog
Dim Lst As ListItem
Dim Msg$
Dim prp
Screen.MousePointer = vbHourglass
Dlink.hWnd = Me.hWnd
Set Obj = Dlink.PromptNew       '---Prompt the user for selecting any database that is _
                                    registered in that computer
Dim tbl As Table
If Obj Is Nothing Then
    Screen.MousePointer = vbDefault
    Exit Sub '---If the user pressed cancel then nothing to do
End If

Msg = "Below is the connection details for the database you have just selected. " & _
      " What you have to do is just pass this connection string while opening using " & Chr(34) & _
      "Open" & Chr(34) & " property of ADO Connect Object."
MsgBox Msg & vbCr & vbCr & Obj, vbInformation
If Conn.State = 1 Then Conn.Close '---Close the already opened global Conn
Conn.Open Obj   '---Open the connection
'---If the conection is success then get the table details of that database
'---Set the catalog connection to just opened one
Catlg.ActiveConnection = Conn
'---Clear the table list
LstTable.ListItems.Clear
'---Get all tables and list it in the Lsi view
For Each tbl In Catlg.Tables
'---Get only tables
    If tbl.Type = "TABLE" Then
        Set Lst = LstTable.ListItems.Add(, , tbl.Name, 1, 1)
        Lst.SubItems(1) = IIf(IsNull(tbl.DateCreated), "", tbl.DateCreated) 'Lst.SubItems(1) = tbl.DateCreated
        Lst.SubItems(2) = IIf(IsNull(tbl.DateModified), "", tbl.DateModified)
        Lst.SubItems(3) = tbl.Type
    End If
Next
LstTable.ColumnHeaders(1).Width = 4000
LstTable.ColumnHeaders(2).Width = 3000
LstTable.ColumnHeaders(3).Width = 3000
'---If successfull retrievel of tables then chenge the Caption to the current connection
TxtConnectionString = Obj
Screen.MousePointer = vbDefault
Exit Sub
errH:
MsgBox "Error Occured while connection." & "[" & Err.Description & ".]", vbCritical
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Set Me.Icon = ImageList1.ListImages(1).Picture
Set MDIForm1.Icon = ImageList1.ListImages(1).Picture
End Sub

Private Sub Form_Resize()
On Error Resume Next
TxtConnectionString.Width = ScaleWidth - TxtConnectionString.Left
LstTable.Left = 0
LstTable.Height = ScaleHeight - LstTable.Top
LstTable.Width = ScaleWidth
End Sub


Private Sub Form_Unload(Cancel As Integer)
'End
End Sub


Private Sub LstTable_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
LstTable.Sorted = True
    LstTable.SortKey = ColumnHeader.Index - 1
    If LstTable.SortOrder = lvwAscending Then
        LstTable.SortOrder = lvwDescending
    Else
        LstTable.SortOrder = lvwAscending
    End If
    LstTable.Sorted = False
    Exit Sub
End Sub

Private Sub LstTable_DblClick()
On Error Resume Next
If LstTable.ListItems.Count = 0 Then Exit Sub   '---If there are no tables then nothing to do
If Conn.State = 0 Then
    If Not Trim$(TxtConnectionString) = "" Then
        Conn.Open Trim$(TxtConnectionString)
        If Err Then
            MsgBox Err.Description, vbExclamation
            LstTable.ListItems.Clear
            Exit Sub
        End If
    End If
End If
        
Dim GetRec As New ADODB.Recordset   '---Pointer to the table the user is selected
Dim Frm As Form                     '---New instance of Entry form
Dim prp As ADODB.Property
Dim i&, j&                          '---For loop
Dim ItemtoAdd$
If OptPick Then
    Set Frm = New FrmPicklist              '---Setting frm to FrmEntry so that everytime new FrmEntry is opened for each table
Else
    Set Frm = New FrmEntry
End If
Frm.Show                            '---Show the form


Frm.Caption = "Table : [" & LstTable.SelectedItem.Text & "];" & Space(5) & " No Records"
Set Frm.Icon = ImageList1.ListImages(1).Picture '---Setting the Forms icon
Frm.SetFocus
Frm.Refresh
Closing = False
If OptPick Then
    Frm.PickList1.PicklistConnection = Conn
'---At first check whether it is a valid Table Name... _
    This is because for oracle database we have to pass Username.Tablename format
    GetRec.Open "SELECT * FROM [" & LstTable.SelectedItem.Text & "]", Conn, adOpenStatic, adLockReadOnly
    If Err Then GoTo errH1
    Frm.PickList1.RetrieveData "SELECT * FROM [" & LstTable.SelectedItem.Text & "]"
    Frm.Caption = "Table : [" & LstTable.SelectedItem.Text & "]" & Space(5) & Frm.PickList1.RecordCount & " Record(s)"
    Frm.PickList1.picklistCaption = "Table : [" & LstTable.SelectedItem.Text & "]" & Space(5) & Frm.PickList1.RecordCount & " Record(s)"
'---If there is an error while opening the table try to concatenate the user name with the table name _
Eg., In oracle u have to pass the table name like Username.TableName
errH1:
        If Err Then
            Err.Clear
'---This loop search for the user name from the properties collection
            For Each prp In Conn.Properties
                If UCase(Trim$(prp.Name)) = "USER ID" Or UCase(Trim$(prp.Name)) = "USER NAME" Then
                    Frm.PickList1.RetrieveData "SELECT * FROM " & prp.Value & "." & LstTable.SelectedItem.Text
                    Exit For
                End If
            Next
        End If
        If Err Then MsgBox "Error Occured.[" & Err.Description & "]", vbExclamation: Err.Clear: Unload Frm: Exit Sub
Else
    With LstTable
        With GetRec
    '---Open the selected table
            If .State = 1 Then .Close
            .Open "SELECT * FROM [" & LstTable.SelectedItem.Text & "]", Conn, adOpenStatic, adLockReadOnly  ', adCmdUnknown
            If Err Then
                Err.Clear
                For Each prp In Conn.Properties
                    If UCase(Trim$(prp.Name)) = "USER ID" Or UCase(Trim$(prp.Name)) = "USER NAME" Then
                        If .State = 1 Then .Close
                        .Open "SELECT * FROM " & prp.Value & "." & LstTable.SelectedItem.Text, Conn, adOpenStatic, adLockReadOnly   ', adCmdUnknown
                        Exit For
                    End If
                Next
            End If
            If Err Then MsgBox "Error Occured.[" & Err.Description & "]", vbExclamation: Err.Clear: Unload Frm: Exit Sub
            If .RecordCount > 0 Then
                With Frm
    '---If there are records then set the rows and columns according to number of fields and _
        Records in the selected table
                    .GridList.Rows = 1  '---First clear rows and cols
                    .GridList.Cols = 1
                    .GridList.Cols = GetRec.Fields.Count + 1
'                    .GridList.Rows = GetRec.RecordCount + 1
    '---Set the field names
                    For i = 0 To GetRec.Fields.Count - 1
                         .GridList.TextMatrix(0, i + 1) = GetRec.Fields(i).Name
                    Next
                    j = 1
    '---Retrive the records and show according to each field
                    Do Until GetRec.EOF
                        DoEvents
                        If Closing Then Closing = False: Exit Sub
                        Frm.Refresh
                        ItemtoAdd = Chr(9)
                        For i = 1 To .GridList.Cols - 1
                            ItemtoAdd = ItemtoAdd & IIf(IsNull(GetRec.Fields(i - 1).Value), "" & Chr(9), GetRec.Fields(i - 1).Value & Chr(9))
                            'If Not IsNull(GetRec.Fields(i - 1).Value) Then .GridList.TextMatrix(j, i) = GetRec.Fields(i - 1).Value
                        Next
                        .GridList.AddItem ItemtoAdd
                        GetRec.MoveNext
                        If Closing Then Closing = False: Exit Sub
                        .Caption = "Table : [" & LstTable.SelectedItem.Text & "]" & Space(5) & "Retrieving..." & j & " Record(s) of " & GetRec.RecordCount & " Record(s)."
                        j = j + 1
                    Loop
                End With
            End If
        End With
    End With
    Frm.Caption = "Table : [" & LstTable.SelectedItem.Text & "]" & Space(5) & IIf(j = 0, j, j - 1) & " Record(s)"
End If
errH:
End Sub
