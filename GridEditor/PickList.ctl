VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.UserControl PickList 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11310
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   PropertyPages   =   "PickList.ctx":0000
   ScaleHeight     =   5625
   ScaleWidth      =   11310
   ToolboxBitmap   =   "PickList.ctx":0027
   Begin VB.Frame QueryFrame 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   600
      TabIndex        =   2
      Top             =   390
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   2370
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3105
         Width           =   705
      End
      Begin VB.CommandButton CmdCAncel1 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3525
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3360
         Width           =   705
      End
      Begin VB.ListBox HeaderList 
         Height          =   1035
         Left            =   5070
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3150
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.ListBox WidthList 
         Height          =   1035
         Left            =   6645
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.ListBox FieldList 
         Height          =   1035
         Left            =   5865
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3135
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.ComboBox QueryCombo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox EditText 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   600
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DatePicker 
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   22740995
         CurrentDate     =   36794
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   2415
         Left            =   315
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   465
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4260
         _Version        =   393216
         Rows            =   10
         Cols            =   5
         AllowBigSelection=   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   90
         TabIndex        =   14
         Top             =   120
         Width           =   840
      End
      Begin VB.Label LblBack 
         BackColor       =   &H8000000D&
         Height          =   255
         Left            =   45
         TabIndex        =   13
         Top             =   120
         Width           =   7485
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8565
      Top             =   2205
   End
   Begin VB.Frame MainFrame 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   -30
      TabIndex        =   4
      Top             =   510
      Width           =   9255
      Begin VB.CommandButton CmdRefresh1 
         Height          =   345
         Left            =   810
         Picture         =   "PickList.ctx":0339
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4275
         Width           =   330
      End
      Begin VB.CommandButton CmdQuery1 
         Height          =   345
         Left            =   1140
         Picture         =   "PickList.ctx":0483
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4275
         Width           =   330
      End
      Begin VB.CommandButton cmdClose1 
         Height          =   345
         Left            =   1470
         Picture         =   "PickList.ctx":0585
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4275
         Width           =   330
      End
      Begin VB.CommandButton CmdSelect1 
         Height          =   345
         Left            =   480
         Picture         =   "PickList.ctx":06CF
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4275
         Width           =   330
      End
      Begin VB.CommandButton CmdSelect 
         Caption         =   "Select"
         Height          =   360
         Left            =   1470
         TabIndex        =   22
         Top             =   4710
         Width           =   795
      End
      Begin VB.CommandButton CmdQuery 
         Caption         =   "Query"
         Height          =   360
         Left            =   3060
         TabIndex        =   21
         Top             =   4710
         Width           =   795
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh"
         Height          =   360
         Left            =   2265
         TabIndex        =   20
         Top             =   4710
         Width           =   795
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         Height          =   360
         Left            =   3855
         TabIndex        =   19
         Top             =   4710
         Width           =   795
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   3150
         TabIndex        =   16
         Top             =   4320
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4080
         Left            =   195
         TabIndex        =   10
         Top             =   615
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   7197
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox SearchText 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   9015
      End
      Begin VB.Label MsgLabel 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   4920
         Width           =   45
      End
   End
   Begin VB.Label Lbl_Caption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   405
      TabIndex        =   18
      Top             =   105
      Width           =   105
   End
   Begin VB.Label LblBackground 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   75
      TabIndex        =   17
      Top             =   75
      Width           =   8850
   End
   Begin VB.Shape Border2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   5565
      Left            =   60
      Top             =   45
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Shape Border1 
      BorderWidth     =   2
      Height          =   5625
      Left            =   30
      Top             =   15
      Visible         =   0   'False
      Width           =   9285
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Men 
         Caption         =   "Select"
         Index           =   0
      End
      Begin VB.Menu Men 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Men 
         Caption         =   "Query"
         Index           =   2
      End
      Begin VB.Menu Men 
         Caption         =   "Refresh"
         Index           =   3
      End
      Begin VB.Menu Men 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu Men 
         Caption         =   "Select All"
         Index           =   5
      End
      Begin VB.Menu Men 
         Caption         =   "Deselect All"
         Index           =   6
      End
      Begin VB.Menu Men 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu Men 
         Caption         =   "Cancel"
         Index           =   8
      End
   End
End
Attribute VB_Name = "PickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'************************************************************************
'*  Programmer   :   SHAHIM.A.C                                         *
'*  Control      :   Picklist Control                                   *
'*  Purpose      :   Retrieves data according to the query passed and   *
'*                   shows in a list view control                       *
'************************************************************************
Option Explicit
Private MainData As New ADODB.Connection
Private CurrentRec1 As New ADODB.Recordset
Private FieldCount As Integer
Private itmFound As ListItem
Private MainQuery, CurrentQuery As String
Private ColumnIndex, ColumnIndex1 As Integer
Private showTools1 As Boolean
Private lastError1 As String
Private ConnectionString1 As String
Private RecordCount1 As Long
Private Refreshing As Boolean
Private CurColumnCount As Long
Private ShowMenu1 As Boolean
Private ShowText1 As Boolean
Private ShowProgress1 As Boolean
Private MultiSelect1 As Boolean
Private CheckBox1 As Boolean
Private IconType1 As IconSize
Private ShowCaption1 As Boolean
Private MovablePicklist As Boolean
Private DateFilterStartFormat1 As String
Private DateFilterEndFormat1 As String
Private DateFormat1 As String

'-----------Event Declarations----------------
Public Event OnSelectClicked()
Public Event OnCancelClicked()
Public Event OnRefreshClicked(Cancel As Boolean)
Public Event OnQueryClicked(Cancel As Boolean)
Public Event ItemClick(ByVal Item As MSComctlLib.ListItem)
Public Event ItemCheck(ByVal Item As MSComctlLib.ListItem)
'public event
'For moving the control
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Public Enum IconSize
    SmallIcons = 1
    LargeIcons = 2
End Enum

Public Sub ClearAll()
On Error Resume Next
    ListView1.ListItems.Clear
End Sub
Private Sub Resize()
On Error Resume Next
'-------Hide ToolButtons
CmdSelect.Visible = False
CmdRefresh.Visible = False
CmdQuery.Visible = False
CmdCancel.Visible = False
CmdSelect1.Visible = False
CmdRefresh1.Visible = False
CmdQuery1.Visible = False
cmdClose1.Visible = False

'-------------Set Caption Label-----------------------------
If ShowCaption1 Then
    Border1.Visible = True
    Border2.Visible = True
Else
    Border1.Visible = False
    Border2.Visible = False
End If
Border1.Width = UserControl.Width - 30
Border2.Width = UserControl.Width - Border2.Left - 30
Border1.Height = UserControl.Height - 60
Border2.Height = UserControl.Height - 120 '(Border2.Top + 60)
With LblBackground
    .Width = UserControl.Width - 150
    Lbl_Caption.Width = .Width
    Lbl_Caption.Left = .Left
End With
'-------------Set command Buttons Left-----------------------------
    CmdQuery.Left = UserControl.Width / 2
    CmdRefresh.Left = CmdQuery.Left - CmdRefresh.Width - 100
    CmdSelect.Left = CmdRefresh.Left - CmdSelect.Width - 100
    CmdCancel.Left = CmdQuery.Left + CmdQuery.Width + 100
    
    CmdSelect1.Left = 0
    CmdRefresh1.Left = CmdSelect1.Left + CmdSelect1.Width '+ 20
    CmdQuery1.Left = CmdRefresh1.Left + CmdRefresh1.Width '+ 20
    cmdClose1.Left = CmdQuery1.Left + CmdQuery1.Width '+ 20
    
'-------------Set command Buttons Top-----------------------------
    CmdQuery.Top = MainFrame.Height - CmdQuery.Height '- 100
    CmdRefresh.Top = CmdQuery.Top
    CmdSelect.Top = CmdQuery.Top
    CmdCancel.Top = CmdQuery.Top
    
    CmdSelect1.Top = 0
    CmdRefresh1.Top = 0
    CmdQuery1.Top = 0
    cmdClose1.Top = 0
    
'---------Set Mainframe
If ShowCaption1 Then
    MainFrame.Top = LblBackground.Height + 100
    MainFrame.Left = LblBackground.Left
    MainFrame.Height = UserControl.Height - (LblBackground.Height + LblBackground.Top + 150)
    MainFrame.Width = LblBackground.Width
Else
    MainFrame.Top = 0
    MainFrame.Left = 0
    MainFrame.Width = UserControl.ScaleWidth
    MainFrame.Height = UserControl.Height
End If
    
    
ListView1.Left = 0
ListView1.Width = MainFrame.Width 'UserControl.ScaleWidth

'---------Set TextBox
SearchText.Left = 0
SearchText.Top = 0
SearchText.Width = MainFrame.Width 'UserControl.ScaleWidth
'--------Set Query Frame
QueryFrame.Top = 0
QueryFrame.Left = 0
QueryFrame.Width = UserControl.ScaleWidth
QueryFrame.Height = UserControl.Height

LblBack.Width = QueryFrame.Width - 80
CmdOk.Top = QueryFrame.Height - CmdOk.Height - 75
CmdCAncel1.Top = CmdOk.Top

'--------Set Query Grid
Grid1.Top = LblBack.Top + LblBack.Height + 10
Grid1.Left = QueryFrame.Left + 75
Grid1.Width = QueryFrame.Width - 150
Grid1.Height = UserControl.Height - LblBack.Height - 250 - CmdOk.Height
    
CmdCAncel1.Left = Grid1.Left + Grid1.Width - CmdCAncel1.Width
CmdOk.Left = CmdCAncel1.Left - CmdOk.Width '- 50

If Not ShowTools And Not ShowText Then
    ListView1.Top = 0
    ListView1.Height = MainFrame.Height
    Exit Sub
End If
If ShowText And Not ShowTools Then
    ListView1.Top = SearchText.Height + 20
    ListView1.Height = MainFrame.Height - SearchText.Height
    Exit Sub
End If
If Not ShowText And ShowTools Then
    If IconType = LargeIcons Then
        ListView1.Top = 0
        ListView1.Height = MainFrame.Height - CmdSelect.Height - 75
        CmdSelect.Visible = True
        CmdRefresh.Visible = True
        CmdQuery.Visible = True
        CmdCancel.Visible = True
    Else
        ListView1.Top = CmdSelect1.Height + 100
        ListView1.Height = MainFrame.Height - CmdSelect1.Height - 100
        CmdSelect1.Visible = True
        CmdRefresh1.Visible = True
        CmdQuery1.Visible = True
        cmdClose1.Visible = True
    End If
Exit Sub
End If
If ShowText And ShowTools Then
    If IconType = LargeIcons Then
        ListView1.Top = SearchText.Height + 20
        ListView1.Height = MainFrame.Height - SearchText.Height - CmdSelect.Height - 75
        CmdSelect.Visible = True
        CmdRefresh.Visible = True
        CmdQuery.Visible = True
        CmdCancel.Visible = True
    Else
        SearchText.Top = CmdSelect1.Top + CmdSelect1.Width + 40
        ListView1.Top = SearchText.Top + SearchText.Height + 20
        ListView1.Height = MainFrame.Height - CmdSelect1.Height - SearchText.Height - 100
        CmdSelect1.Visible = True
        CmdRefresh1.Visible = True
        CmdQuery1.Visible = True
        cmdClose1.Visible = True
    End If
End If
End Sub

Private Sub Resizeold()
On Error Resume Next
'-------Hide ToolButtons
CmdSelect.Visible = False
CmdRefresh.Visible = False
CmdQuery.Visible = False
CmdCancel.Visible = False
CmdSelect1.Visible = False
CmdRefresh1.Visible = False
CmdQuery1.Visible = False
cmdClose1.Visible = False

'-------------Set command Buttons Left-----------------------------
    CmdQuery.Left = UserControl.Width / 2
    CmdRefresh.Left = CmdQuery.Left - CmdRefresh.Width - 100
    CmdSelect.Left = CmdRefresh.Left - CmdSelect.Width - 100
    CmdCancel.Left = CmdQuery.Left + CmdQuery.Width + 100
    
    CmdSelect1.Left = 0
    CmdRefresh1.Left = CmdSelect1.Left + CmdSelect1.Width '+ 20
    CmdQuery1.Left = CmdRefresh1.Left + CmdRefresh1.Width '+ 20
    cmdClose1.Left = CmdQuery1.Left + CmdQuery1.Width '+ 20
    
'-------------Set command Buttons Top-----------------------------
    CmdQuery.Top = UserControl.Height - CmdQuery.Height '- 100
    CmdRefresh.Top = CmdQuery.Top
    CmdSelect.Top = CmdQuery.Top
    CmdCancel.Top = CmdQuery.Top
    
    CmdSelect1.Top = 0
    CmdRefresh1.Top = 0
    CmdQuery1.Top = 0
    cmdClose1.Top = 0
    
'--------Set Borders
'Border1.Left = 0
'Border2.Left = 20
'Border1.Top = 0
'Border2.Top = 20
'Border1.Width = UserControl.Width
'Border2.Width = UserControl.Width - 40
'Border1.Height = UserControl.Height
'Border2.Height = UserControl.Height - 40


'---------Set Mainframe

MainFrame.Top = 0
MainFrame.Left = 0
MainFrame.Width = UserControl.ScaleWidth
MainFrame.Height = UserControl.ScaleHeight

ListView1.Left = 0
ListView1.Width = UserControl.ScaleWidth

'---------Set TextBox
SearchText.Left = 0
SearchText.Top = 0
SearchText.Width = UserControl.ScaleWidth
'--------Set Query Frame
QueryFrame.Top = 0
QueryFrame.Left = 0
QueryFrame.Width = UserControl.ScaleWidth
QueryFrame.Height = UserControl.ScaleHeight

LblBack.Width = QueryFrame.Width - 80
CmdOk.Top = QueryFrame.Height - CmdOk.Height - 75
CmdCAncel1.Top = CmdOk.Top

'--------Set Query Grid
Grid1.Top = LblBack.Top + LblBack.Height + 10
Grid1.Left = QueryFrame.Left + 75
Grid1.Width = QueryFrame.Width - 150
Grid1.Height = UserControl.ScaleHeight - LblBack.Height - 250 - CmdOk.Height
    
CmdCAncel1.Left = Grid1.Left + Grid1.Width - CmdCAncel1.Width
CmdOk.Left = CmdCAncel1.Left - CmdOk.Width '- 50

If Not ShowTools And Not ShowText Then
    ListView1.Top = 0
    ListView1.Height = UserControl.ScaleHeight
    Exit Sub
End If
If ShowText And Not ShowTools Then
    ListView1.Top = SearchText.Height + 20
    ListView1.Height = UserControl.ScaleHeight - SearchText.Height
    Exit Sub
End If
If Not ShowText And ShowTools Then
    If IconType = LargeIcons Then
        ListView1.Top = 0
        ListView1.Height = UserControl.ScaleHeight - CmdSelect.Height - 75
        CmdSelect.Visible = True
        CmdRefresh.Visible = True
        CmdQuery.Visible = True
        CmdCancel.Visible = True
    Else
        ListView1.Top = CmdSelect1.Height + 100
        ListView1.Height = UserControl.ScaleHeight - CmdSelect1.Height - 100
        CmdSelect1.Visible = True
        CmdRefresh1.Visible = True
        CmdQuery1.Visible = True
        cmdClose1.Visible = True
    End If
Exit Sub
End If
If ShowText And ShowTools Then
    If IconType = LargeIcons Then
        ListView1.Top = SearchText.Height + 20
        ListView1.Height = UserControl.ScaleHeight - SearchText.Height - CmdSelect.Height - 75
        CmdSelect.Visible = True
        CmdRefresh.Visible = True
        CmdQuery.Visible = True
        CmdCancel.Visible = True
    Else
        SearchText.Top = CmdSelect1.Top + CmdSelect1.Width + 40
        ListView1.Top = SearchText.Top + SearchText.Height + 20
        ListView1.Height = UserControl.ScaleHeight - CmdSelect1.Height - SearchText.Height - 100
        CmdSelect1.Visible = True
        CmdRefresh1.Visible = True
        CmdQuery1.Visible = True
        cmdClose1.Visible = True
    End If
End If
End Sub


Private Sub SetUpProgressBar(Count1 As Long)
On Error Resume Next
    ProgressBar1.Max = Count1
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0
    ProgressBar1.Width = ListView1.Width - 30
    ProgressBar1.Top = ListView1.Top + ListView1.Height - ProgressBar1.Height - 50
    ProgressBar1.Left = ListView1.Left + 10
    ProgressBar1.Visible = True
    UserControl.Refresh
End Sub
Public Sub RefreshData()
    RetrieveData MainQuery
End Sub

Private Function Retrieve(Optional ByVal QueryString As String) As Integer
On Error GoTo errorHand
Dim i, j As Long
Dim RecCount As Long
Dim GetRec As New ADODB.Recordset
Retrieve = 0
    With CurrentRec1
        Debug.Print QueryString
        .Filter = QueryString
        If .RecordCount > 0 Then
            .MoveLast
            RecCount = .RecordCount
            If ShowProgress1 Then SetUpProgressBar RecCount
            .MoveFirst
        Else
        'MsgBox "No Matching Records Found", vbExclamation
            Retrieve = -1
            Exit Function
        End If
        
        ListView1.ListItems.Clear                               'Clear List view
        ListView1.ColumnHeaders.Clear                           'Clear Column Headers

        For i = 0 To .Fields.Count - 1
            ListView1.ColumnHeaders.Add , , .Fields(i).Name     'Show fieldnames on ColumnHeaders
            FieldList.AddItem .Fields(i).Name
        Next
        
        For i = 1 To RecCount   'Fill ListView
        'First add the 1st column the go for subitems of that row
            If IsNull(.Fields(0)) Then  'Add the first value.If null then add nothing
                ListView1.ListItems.Add , , ""
            Else
                ListView1.ListItems.Add , , .Fields(0).Value
            End If
            For j = 1 To .Fields.Count - 1
                If IsNull(.Fields(j)) Then  'Add rest of the fields(SubItems)
                    ListView1.ListItems.Item(i).SubItems(j) = ""
                Else
                    ListView1.ListItems.Item(i).SubItems(j) = .Fields(j).Value
                End If
            Next
            .MoveNext
            UpdateProgress Val(i)
        Next
    End With 'End of GetRec
    On Error Resume Next
    With WidthList
            For i = 0 To .ListCount - 1
            ListView1.ColumnHeaders(Val(Left(.List(i), InStr(1, .List(i), ";") - 1))).Width = Right(.List(i), Len(.List(i)) - InStr(1, .List(i), ";"))
        Next
    End With
    With HeaderList
            For i = 0 To .ListCount - 1
                ListView1.ColumnHeaders(Val(Left(.List(i), InStr(1, .List(i), ";") - 1))).Text = Right(.List(i), Len(.List(i)) - InStr(1, .List(i), ";"))
            Next
        End With
    ProgressBar1.Visible = False
Exit Function
errorHand:
Retrieve = 1
ProgressBar1.Visible = False
End Function

Public Sub RetrieveData(ByVal QueryString As String)
Dim i, j As Long
On Error GoTo errHand
MainQuery = QueryString
CurColumnCount = 0
FieldList.Clear
If Trim$(MainQuery) = "" Then
    Err.Raise -90900, , "No Query string passed"
    Exit Sub
End If
Dim GetRec As New ADODB.Recordset
ListView1.ListItems.Clear                               'Clear List view
ListView1.ColumnHeaders.Clear                           'Clear Column Headers
RecordCount1 = 0
    If MainData Is Nothing Then 'Check for maindata Connection. If nothing then raise error
        Err.Raise 1000034, MainData, "Picklist is not connected to any Database/Connection."
        Exit Sub
    End If
    
    With GetRec
        .Open QueryString, MainData, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            RecordCount1 = .RecordCount
            .MoveFirst
            If ShowProgress1 Then Call SetUpProgressBar(RecordCount1)
        Else
            Exit Sub
        End If
        CurColumnCount = .Fields.Count
'---Show fieldnames on ColumnHeaders
        For i = 0 To .Fields.Count - 1
            ListView1.ColumnHeaders.Add , , .Fields(i).Name
            FieldList.AddItem .Fields(i).Name
        Next
        
        For i = 1 To RecordCount1   'Fill ListView
'---First add the 1st column the go for subitems of that row
            If IsNull(.Fields(0)) Then  'Add the first value.If null then add nothing
                ListView1.ListItems.Add , , ""
            Else
                ListView1.ListItems.Add , , .Fields(0).Value
            End If
'---Add rest of the fields(SubItems)
            For j = 1 To .Fields.Count - 1
                If IsNull(.Fields(j)) Then
                    ListView1.ListItems.Item(i).SubItems(j) = ""
                Else
                    ListView1.ListItems.Item(i).SubItems(j) = .Fields(j).Value
                End If
'                DoEvents
            Next
            .MoveNext
            If ShowProgress1 Then Call UpdateProgress(Val(i))
        Next
        On Error Resume Next
'---Set the column width which is stored in 'WidthList' Listbox
        With WidthList
            For i = 0 To .ListCount - 1
                ListView1.ColumnHeaders(Val(Left(.List(i), InStr(1, .List(i), ";") - 1))).Width = Right(.List(i), Len(.List(i)) - InStr(1, .List(i), ";"))
            Next
        End With
'---Set the header names
        With HeaderList
            For i = 0 To .ListCount - 1
                ListView1.ColumnHeaders(Val(Left(.List(i), InStr(1, .List(i), ";") - 1))).Text = Right(.List(i), Len(.List(i)) - InStr(1, .List(i), ";"))
            Next
        End With
    End With 'End of GetRec
    ProgressBar1.Visible = False
    Exit Sub
errHand:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    ProgressBar1.Visible = False
End Sub

Public Sub ShowAbout()
Attribute ShowAbout.VB_UserMemId = -552
'    Frm_About_Picklist.Show vbModal
End Sub
Private Function getFieldName(Caption As String) As String
On Error Resume Next
    Dim i As Long
        For i = 1 To ListView1.ColumnHeaders.Count
            If ListView1.ColumnHeaders(i).Text = Caption Then
                getFieldName = FieldList.List(i - 1)
                Exit For
            End If
        Next
End Function

Private Sub Hidecontrols()
On Error Resume Next
'This function is used to hide all editing controls/objects from the query grid
    QueryCombo.Visible = False
    EditText.Visible = False
    DatePicker.Visible = False
End Sub

Public Sub RemoveItem(ByVal Index As Integer)
'Sub used to remove a row from the listview
On Error Resume Next
    ListView1.ListItems.Remove Index
End Sub
Private Sub SetControlonGrid()
'This sub inserts appropriate object into the query grid
'This function gets the field type using getFieldType function and according to
'the type it insets textbox or date control box
On Error Resume Next
    Call Hidecontrols
    Select Case getFieldType
        Case adDate, adDBDate, 135 'For date
            Call SetDateControl
        Case Else
            Call SetTextBox
    End Select
 lastError1 = Err.Description:
End Sub

Private Sub UpdateProgress(Value1 As Long)
On Error Resume Next
ProgressBar1.Value = Value1
End Sub

Private Sub CmdCancel_Click()
On Error Resume Next
    SearchText = ""
    RaiseEvent OnCancelClicked
lastError1 = Err.Description: End Sub

Private Sub CmdCAncel1_Click()
On Error Resume Next
    QueryFrame.Visible = False
    MainFrame.Enabled = True
    SearchText.SetFocus
    CmdOk.Default = False
    CmdSelect.Default = True
lastError1 = Err.Description:
End Sub

Private Sub cmdClose1_Click()
    Call CmdCancel_Click
End Sub

Private Sub CmdOk_Click()
On Error Resume Next
QueryFrame.Visible = False: MainFrame.Enabled = True: SearchText.SetFocus
Call CreateQuery
    If CurrentQuery = "" Then QueryFrame.Visible = False: MainFrame.Enabled = True: Exit Sub
    Select Case Retrieve(CurrentQuery)
        Case 1, -1
            MsgBox "No Matching Records Found.", vbInformation, "Query"
    End Select
    CmdOk.Default = False
    CmdSelect.Default = True
    ProgressBar1.Visible = False
    SearchText = ""
    SearchText.SetFocus
lastError1 = Err.Description
End Sub

Private Function getFieldType(Optional ByVal Gridrow As Long) As Integer
'This function returns the fieldtype of a particular field
On Error Resume Next
Dim GridRow1 As Long
If Gridrow = 0 Then
GridRow1 = Grid1.Row
Else
GridRow1 = Gridrow
End If
With Grid1
    getFieldType = CurrentRec1.Fields(getFieldName(.TextMatrix(.Row, 1))).Type 'Grid1.Row - 1)
End With
lastError1 = Err.Description:
End Function

Private Sub SetCombo()
On Error Resume Next
    With QueryCombo
        .Top = Grid1.Top + Grid1.CellTop
        .Left = Grid1.Left + Grid1.CellLeft
        .Width = Grid1.CellWidth
        .Visible = True
        .SetFocus
        Call ComboWithOperators
    End With
    
lastError1 = Err.Description: End Sub
Private Sub SetTextBox()
On Error Resume Next
    With EditText
        .Top = Grid1.Top + Grid1.CellTop
        .Left = Grid1.Left + Grid1.CellLeft
        .Width = Grid1.CellWidth
        .Height = Grid1.CellHeight
        .Visible = True
        .SetFocus
    End With
    
lastError1 = Err.Description: End Sub
Private Sub ComboWithOperators()
Dim i As Long
On Error Resume Next
    With QueryCombo
    .Clear
       If Grid1.Col = 2 Then
        Select Case getFieldType(Grid1.Row)
            Case adChar, adVarChar, adLongVarChar, adLongVarWChar, 202
                .AddItem "="
                .AddItem "Like"
            Case Else
                .AddItem "="
                .AddItem ">"
                .AddItem "<"
                .AddItem ">="
                .AddItem "<="
                .AddItem "<>"
                .AddItem "Like"
         End Select
      ElseIf Grid1.Col = 4 Then
        .AddItem "And"
        .AddItem "Or"
      ElseIf Grid1.Col = 1 Then
        With Grid1
            For i = 1 To ListView1.ColumnHeaders.Count
                QueryCombo.AddItem ListView1.ColumnHeaders(i).Text
            Next
        End With
       End If
    End With
lastError1 = Err.Description: End Sub
Private Sub GridSerial()
Dim i As Long
On Error Resume Next
    For i = 1 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 0) = i
    Next i
lastError1 = Err.Description: End Sub

Private Function CreateQuery()
Dim i As Long
Dim DtSt$, DtEnd$
DtSt = DateFilterStartFormat1
DtEnd = DateFilterEndFormat1
'This function creates query, when the user clicks on query button
On Error Resume Next
    Dim temp As String
        For i = 1 To Grid1.Rows - 1
            Select Case getFieldType(i)
                   Case 3, adInteger, adSingle, adSmallInt, adNumeric, adDouble, adCurrency ' 16, 20, 21, 3, 4, 19, 6   'For NUmbers
                    If Grid1.TextMatrix(i, 2) = "" Or Grid1.TextMatrix(i, 3) = "" Then GoTo skip
                        If Grid1.TextMatrix(i, 2) = "Like" Then
                            temp = temp & "Str([" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "])" & " Like '" & Trim$(Grid1.TextMatrix(i, 3)) & "'" & " " & Grid1.TextMatrix(i, 4) & " "
                        Else
                            temp = temp & " [" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "] " & Grid1.TextMatrix(i, 2) & " " & Grid1.TextMatrix(i, 3) & " " & Grid1.TextMatrix(i, 4) & " "
                        End If
                   Case adChar, adVarChar, adLongVarChar, adLongVarWChar, 202 ', 10, 12, 18, 202 'For Varchars
                   If Grid1.TextMatrix(i, 2) = "" Or Grid1.TextMatrix(i, 3) = "" Then GoTo skip
                        If Grid1.TextMatrix(i, 2) = "Like" Then
                            temp = temp & " [" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "] " & " Like '" & Grid1.TextMatrix(i, 3) & "'" & " " & Grid1.TextMatrix(i, 4) & " "
                        Else
                            temp = temp & " [" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "] " & Grid1.TextMatrix(i, 2) & " '" & Grid1.TextMatrix(i, 3) & "'" & " " & Grid1.TextMatrix(i, 4) & " "
                        End If
                   Case adDate, adDBDate, 135
                   If Grid1.TextMatrix(i, 2) = "" Or Grid1.TextMatrix(i, 3) = "" Then GoTo skip
                        If Grid1.TextMatrix(i, 2) = "Like" Then
                            temp = temp & " [" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "] " & " Like " & DtSt & Format(Grid1.TextMatrix(i, 3), DateFormat1) & DtEnd & " " & Grid1.TextMatrix(i, 4) & " "
                        Else
                            temp = temp & " [" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "] " & Grid1.TextMatrix(i, 2) & DtSt & Format(Grid1.TextMatrix(i, 3), DateFormat1) & DtEnd & Grid1.TextMatrix(i, 4) & " "
                        End If
            End Select
skip:
        Next i
    If Right$(temp, 4) = "And " Then
        temp = Left$(temp, Len(temp) - 4)
    End If

    If Right$(temp, 3) = "Or " Then
        temp = Left$(temp, Len(temp) - 3)
    End If

    If Right$(temp, 6) = "where " Then
        temp = Left$(temp, Len(temp) - 6)
    End If
    
    CurrentQuery = temp
    MsgBox temp
lastError1 = Err.Description:
End Function


Private Function CreateQuery1()
Dim i As Long
'This function creates query, when the user clicks on query button
On Error Resume Next
    Dim temp As String
        For i = 1 To Grid1.Rows - 1
            Select Case getFieldType(i)
                   Case 3, adInteger, adSingle, adSmallInt, adNumeric, adDouble, adCurrency ' 16, 20, 21, 3, 4, 19, 6   'For NUmbers
                    If Grid1.TextMatrix(i, 2) = "" Or Grid1.TextMatrix(i, 3) = "" Then GoTo skip
                        If Grid1.TextMatrix(i, 2) = "Like" Then
                            temp = temp & "Str([" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "])" & " Like '" & Trim$(Grid1.TextMatrix(i, 3)) & "%'" & " " & Grid1.TextMatrix(i, 4) & " "
                        Else
                            temp = temp & " [" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "] " & Grid1.TextMatrix(i, 2) & " " & Grid1.TextMatrix(i, 3) & " " & Grid1.TextMatrix(i, 4) & " "
                        End If
                   Case adChar, adVarChar, adLongVarChar, adLongVarWChar, 202 ', 10, 12, 18, 202 'For Varchars
                   If Grid1.TextMatrix(i, 2) = "" Or Grid1.TextMatrix(i, 3) = "" Then GoTo skip
                        If Grid1.TextMatrix(i, 2) = "Like" Then
                            temp = temp & " [" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "] " & " Like '" & Grid1.TextMatrix(i, 3) & "%'" & " " & Grid1.TextMatrix(i, 4) & " "
                        Else
                            temp = temp & " [" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "] " & Grid1.TextMatrix(i, 2) & " '" & Grid1.TextMatrix(i, 3) & "'" & " " & Grid1.TextMatrix(i, 4) & " "
                        End If
                   Case adDate, adDBDate
                   If Grid1.TextMatrix(i, 2) = "" Or Grid1.TextMatrix(i, 3) = "" Then GoTo skip
                        If Grid1.TextMatrix(i, 2) = "Like" Then
                            temp = temp & " [" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "] " & " Like #" & CDate(Format(Grid1.TextMatrix(i, 3), "mm/dd/yyyy")) & "%#" & " " & Grid1.TextMatrix(i, 4) & " "
                        Else
                            temp = temp & " [" & getFieldName(Trim(Grid1.TextMatrix(i, 1))) & "] " & Grid1.TextMatrix(i, 2) & "#" & CDate(Format(Grid1.TextMatrix(i, 3), "mm/dd/yyyy")) & "# " & Grid1.TextMatrix(i, 4) & " "
                        End If
            End Select
skip:
        Next i
    If Right$(temp, 4) = "And " Then
        temp = Left$(temp, Len(temp) - 4)
    End If

    If Right$(temp, 3) = "Or " Then
        temp = Left$(temp, Len(temp) - 3)
    End If

    If Right$(temp, 6) = "where " Then
        temp = Left$(temp, Len(temp) - 6)
    End If
    
    CurrentQuery = temp
'    MsgBox temp
lastError1 = Err.Description:
End Function

Private Sub CmdQuery_Click()
On Error Resume Next
Dim Cancel As Boolean
RaiseEvent OnQueryClicked(Cancel)
If Cancel Then Exit Sub
Dim i As Integer
If MainQuery = "" Then Exit Sub
    
    With Grid1
        .Clear
        .TextMatrix(0, 0) = "#"
        .TextMatrix(0, 1) = "Column Name"
        .TextMatrix(0, 2) = "Operator"
        .TextMatrix(0, 3) = "Value"
        For i = 0 To .Rows - 1
            .RowHeight(i) = QueryCombo.Height
        Next
        .Row = 0
        For i = 1 To .Cols - 1
            .Col = i
            .CellAlignment = 4
            .CellFontBold = True
        Next
        .RowHeight(0) = 350
        .ColWidth(4) = 500
    End With
Call Hidecontrols
QueryFrame.Visible = True
For i = 1 To ListView1.ColumnHeaders.Count
    QueryCombo.AddItem ListView1.ColumnHeaders(i).Text
Next
If CurrentRec1.State = 1 Then CurrentRec1.Close
CurrentRec1.Open MainQuery, MainData
SearchText = ""
MainFrame.Enabled = False
CmdOk.Default = True
CmdCAncel1.Cancel = True
CmdSelect.Default = False
Grid1.Col = 3: Grid1.Col = 1:  Grid1.Row = 1
Grid1.SetFocus
lastError1 = Err.Description
End Sub
 
Private Sub CmdQuery1_Click()
CmdQuery_Click
End Sub

Private Sub CmdRefresh_Click()
On Error Resume Next
Dim Cancel As Boolean
RaiseEvent OnRefreshClicked(Cancel)
If Cancel Then Exit Sub
    Refreshing = True
    Call RetrieveData(MainQuery)
    SearchText = ""
    SearchText.SetFocus
lastError1 = Err.Description: End Sub

Private Sub CmdRefresh1_Click()
    Call CmdRefresh_Click
End Sub

Private Sub CmdSelect_Click()
    On Error Resume Next
    'UserControl.Cls
    RaiseEvent OnSelectClicked
    EditText = ""
    SearchText = ""
lastError1 = Err.Description:
End Sub

Private Sub CmdSelect1_Click()
Call CmdSelect_Click
End Sub

Private Sub DatePicker_Change()
On Error Resume Next
Grid1 = DatePicker.Value
With Grid1
        If .Row = .Rows - 1 Then
        If Not .TextMatrix(.Row, 1) = "" Or Not .TextMatrix(.Row, 2) = "" Or Not .TextMatrix(.Row, 3) = "" Or Not .TextMatrix(.Row, 4) = "" Then
            .AddItem ""
            End If
        End If
    End With
lastError1 = Err.Description: End Sub

Private Sub Grid1_EnterCell()
    Select Case Grid1.Col
        Case 1
        Case 2
        Case 3
        Case 4
    End Select
End Sub

Private Sub Grid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Grid1
    Select Case .MouseCol
        Case 1
            .ToolTipText = "Click to select the column you want to check"
        Case 2
            .ToolTipText = "Click to select the operator from the list"
        Case 3
            .ToolTipText = "Click here to enter the search criteria"
        Case 4
            .ToolTipText = "Click here to select the Logical operator And/Or"
        Case Else
            .ToolTipText = ""
    End Select
End With
End Sub

Private Sub Grid1_Scroll()
On Error Resume Next
    DatePicker.Visible = False
    QueryCombo.Visible = False
    EditText.Visible = False
lastError1 = Err.Description: End Sub

Private Sub Lbl_Caption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And MovablePicklist Then
    ReleaseCapture
    Call SendMessage(UserControl.hwnd, &HA1, 2, 0&)
End If
End Sub

Private Sub LblBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And MovablePicklist Then
    ReleaseCapture
    Call SendMessage(UserControl.hwnd, &HA1, 2, 0&)
End If
End Sub

Private Sub ListView1_Click()
On Error Resume Next
    SearchText = ListView1.SelectedItem.Text
lastError1 = Err.Description: End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error Resume Next
   ' Msg "Sorting.."
    ListView1.Sorted = True
    ListView1.SortKey = ColumnHeader.Index - 1
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
    ListView1.Sorted = False
    Exit Sub
lastError1 = Err.Description: End Sub

Private Sub ListView1_DblClick()
    Call CmdSelect_Click
lastError1 = Err.Description: End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    RaiseEvent ItemCheck(Item)
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RaiseEvent ItemClick(Item)
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call CmdSelect_Click
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbRightButton And ShowMenu1 Then
'With Frm_About_Picklist
    If RecordCount1 > 0 Then
        Men(0).Enabled = True
        Men(5).Enabled = True
        Men(6).Enabled = True
    Else
        Men(0).Enabled = False
        Men(5).Enabled = False
        Men(6).Enabled = False
        GoTo skip
        Exit Sub
    End If
    If Not MultiSelect Then
        Men(5).Enabled = False
        Men(6).Enabled = False
        GoTo skip
    End If
    If SelectedCount = 0 Then
        Men(5).Enabled = True
        Men(6).Enabled = False
        GoTo skip
    End If
    If SelectedCount > 0 And Not RowCount = SelectedCount Then
        Men(5).Enabled = True
        Men(6).Enabled = True
        GoTo skip
    End If
    If SelectedCount > 0 And RowCount = SelectedCount Then
        Men(5).Enabled = False
        Men(6).Enabled = True
        GoTo skip
    End If

skip:
    UserControl.PopupMenu Menu1
'    End With
End If
End Sub

Private Sub men_Click(Index As Integer)
On Error Resume Next
Dim i As Long
With ListView1
Select Case Index
    Case 0
        Call CmdSelect_Click
    Case 2
        Call CmdQuery_Click
    Case 3
        RefreshData
    Case 5
        For i = 1 To .ListItems.Count
            .ListItems.Item(i).Selected = True
        Next
    Case 6
        For i = 1 To .ListItems.Count
            .ListItems.Item(i).Selected = False
        Next
    Case 8
        Call CmdCancel_Click
End Select
End With
End Sub


Private Sub SearchText_Change()
On Error Resume Next
   Set itmFound = ListView1. _
   FindItem(SearchText, lvwText, lvwSubItem, lvwPartial)
   If itmFound Is Nothing Then Exit Sub
   itmFound.EnsureVisible
   itmFound.Selected = True
lastError1 = Err.Description: End Sub

Private Sub SearchText_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
    ListView1.SetFocus
ElseIf KeyCode = 13 Then
    Call CmdSelect_Click
End If
lastError1 = Err.Description: End Sub



Private Sub Timer1_Timer()
Call Resize
Timer1.Enabled = False
End Sub

Private Sub UserControl_GotFocus()
On Error Resume Next
If SearchText.Visible Then
    SearchText.SetFocus
Else
    ListView1.SetFocus
End If
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    CurrentRec1.ActiveConnection = MainData
    CurrentRec1.CursorType = adOpenStatic
    CurrentRec1.LockType = adLockReadOnly
    Call SetUpGrid
    ColumnIndex1 = 1
    showTools1 = True
    ShowText1 = True
    ShowMenu1 = True
    ShowProgress1 = True
    IconType1 = LargeIcons
lastError1 = Err.Description: End Sub
Private Sub SetUpGrid()
On Error Resume Next
    With Grid1
        .TextMatrix(0, 0) = "#"
        .TextMatrix(0, 1) = "Column Name"
        .TextMatrix(0, 2) = "Operator"
        .TextMatrix(0, 3) = "Value"

        
        .ColWidth(0) = 300
        .ColWidth(1) = .ColWidth(1) + 750
        .ColWidth(3) = .ColWidth(3) + 1000
        Call GridSerial
    End With
lastError1 = Err.Description: End Sub
Private Sub Grid1_Click()
On Error Resume Next
    If Grid1.Col = 1 Or Grid1.Col = 2 Or Grid1.Col = 4 Then 'Or Grid1.Col = 3 Then
        SetCombo
    If Grid1.Col = 4 And Grid1.Row = Grid1.Rows - 1 Then QueryCombo.Visible = False
        EditText.Visible = False
        DatePicker.Visible = False
    Else:     QueryCombo.Visible = False
    End If
    If Grid1.Col = 3 Then
        Call SetControlonGrid
    End If
    EditText = ""
lastError1 = Err.Description: End Sub
Private Sub SetDateControl()
On Error Resume Next
With DatePicker
        .Top = Grid1.Top + Grid1.CellTop
        .Left = Grid1.Left + Grid1.CellLeft
        .Width = Grid1.CellWidth
        .Height = Grid1.CellHeight
        .Visible = True
        .SetFocus
        Grid1 = DatePicker.Value
End With
lastError1 = Err.Description: End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case Grid1.Col
    Case 1, 2, 4
        SetCombo
        EditText.Visible = False
        DatePicker.Visible = False
End Select

    Select Case KeyAscii
        Case 13, 8, vbKeyTab, vbKeyEscape
        Case Else
            If Grid1.Col <> 3 Then Exit Sub
            EditText = ""
            SetControlonGrid
            EditText = Chr$(KeyAscii)
            EditText.SelStart = Len(EditText)
    End Select
lastError1 = Err.Description: End Sub

Private Sub Grid1_RowColChange()
On Error Resume Next
    Grid1.RowHeight(Grid1.Row) = QueryCombo.Height
    If Grid1.Row <> 1 Then
        If Grid1.TextMatrix(Grid1.Row - 1, 4) = "" Then
            Grid1.TextMatrix(Grid1.Row - 1, 4) = "And"
        End If
    End If
lastError1 = Err.Description: End Sub

Private Sub QueryCombo_Change()
On Error Resume Next
    Grid1.Text = QueryCombo
    With Grid1
        If .Row = .Rows - 1 Then
            If Not .TextMatrix(.Row, 1) = "" Or Not .TextMatrix(.Row, 2) = "" Or Not .TextMatrix(.Row, 3) = "" Or Not .TextMatrix(.Row, 4) = "" Then
                .AddItem ""
            End If
        End If
    End With
lastError1 = Err.Description: End Sub

Private Sub QueryCombo_Click()
On Error Resume Next
 Grid1.Text = QueryCombo
lastError1 = Err.Description: End Sub

Private Sub QueryCombo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyRight Then
        QueryCombo.Visible = False
        If Grid1.Col = 4 Then
            If Grid1.Row <> Grid1.Rows - 1 Then
                Grid1.Row = Grid1.Row + 1
                Grid1.Col = 2
            End If
        ElseIf Grid1.Col = 2 Then
               Grid1.Col = 3
        End If
           Grid1.SetFocus
    End If
lastError1 = Err.Description: End Sub

Private Sub DatePicker_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case 13
            DatePicker.Visible = False
            Grid1 = DatePicker.Value
            Grid1.Col = 4
            Grid1.SetFocus
    End Select
lastError1 = Err.Description: End Sub

Private Sub EditText_Change()
On Error Resume Next
    Grid1 = EditText
    With Grid1
        If .Row = .Rows - 1 Then
            If Not .TextMatrix(.Row, 1) = "" Or Not .TextMatrix(.Row, 2) = "" Or Not .TextMatrix(.Row, 3) = "" Or Not .TextMatrix(.Row, 4) = "" Then
                .AddItem ""
            End If
        End If
    End With
lastError1 = Err.Description: End Sub

Private Sub EditText_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case 13
            EditText.Visible = False
            Grid1 = EditText
            Grid1.Col = 4
            Grid1.SetFocus
    End Select
lastError1 = Err.Description: End Sub
Public Property Get PicklistConnection() As ADODB.Connection
Attribute PicklistConnection.VB_MemberFlags = "400"
lastError1 = Err.Description: End Property

Public Property Let PicklistConnection(ByVal vNewValue As ADODB.Connection)
On Error GoTo errHand
      If MainData.State = 1 Then MainData.Close
        MainData.ConnectionString = vNewValue.ConnectionString
        MainData.Open
    Exit Property
errHand:
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
lastError1 = Err.Description: End Property
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyEscape And QueryFrame.Visible Then
    Call CmdCAncel1_Click
End If
lastError1 = Err.Description
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
   ConnectionString1 = PropBag.ReadProperty("ConnectionString1")
   showTools1 = PropBag.ReadProperty("SHOWTOOL", True)
   ListView1.CheckBoxes = PropBag.ReadProperty("CHECKBOX", ListView1.CheckBoxes)
   ListView1.MultiSelect = PropBag.ReadProperty("MULTISELECT", ListView1.MultiSelect)
   MainQuery = PropBag.ReadProperty("MainQuery")
   ShowMenu1 = PropBag.ReadProperty("SHOWMENU", True)
   Set ListView1.Font = PropBag.ReadProperty("FONT12")
   ListView1.ForeColor = PropBag.ReadProperty("FORECOLOR")
   ShowText1 = PropBag.ReadProperty("SHOWTEXT", True)
   ShowProgress1 = PropBag.ReadProperty("ShowProgress", True)
   IconType1 = PropBag.ReadProperty("IconType1", 2)
   ShowCaption1 = PropBag.ReadProperty("ShowCaption1", True)
   Lbl_Caption.Caption = PropBag.ReadProperty("LBLCAPTION", "")
   Lbl_Caption.BackColor = PropBag.ReadProperty("CAPBACKCOLOR")
   Lbl_Caption.ForeColor = PropBag.ReadProperty("CAPFORECOLOR")
   Border1.BorderColor = PropBag.ReadProperty("BORDERCOLOR1")
   Border2.BorderColor = PropBag.ReadProperty("BORDERCOLOR2")
   Lbl_Caption.Alignment = PropBag.ReadProperty("CAPALIGN")
   Set Lbl_Caption.Font = PropBag.ReadProperty("CAPFONT")
   MovablePicklist = PropBag.ReadProperty("MovablePicklist", False)
   LblBackground.BackColor = PropBag.ReadProperty("LblBackground")
   DateFilterStartFormat1 = PropBag.ReadProperty("DateFilterStartFormat", "#")
   DateFilterEndFormat1 = PropBag.ReadProperty("DateFilterEndFormat1", "#")
   DateFormat1 = PropBag.ReadProperty("DateFormat1", "MM/dd/yyyy")
   'Call SetControl(showTools1)
   SearchText.Visible = ShowText1

If ShowCaption1 Then
    Border1.Visible = True
    Border2.Visible = True
    Lbl_Caption.Visible = True
    LblBackground.Visible = True
Else
    Border1.Visible = False
    Border2.Visible = False
    Lbl_Caption.Visible = False
    LblBackground.Visible = False

End If
   Call UserControl_Resize
lastError1 = Err.Description: End Sub

Private Sub UserControl_Resize()
Call Resize
Timer1.Enabled = True
lastError1 = Err.Description:
End Sub

Public Property Get SetColumnName(ByVal Index As Long) As Variant
Attribute SetColumnName.VB_MemberFlags = "400"
On Error Resume Next
    SetColumnName = ListView1.ColumnHeaders(Index).Text
lastError1 = Err.Description: End Property

Public Property Let SetColumnName(ByVal Index As Long, ByVal vNewValue As Variant)
Dim i As Long
   On Error Resume Next
    ListView1.ColumnHeaders(Index).Text = vNewValue
    For i = 0 To HeaderList.ListCount - 1
    If InStr(1, HeaderList.List(i), Index & ";") > 0 Then
        HeaderList.List(i) = Index & ";" & vNewValue
        Exit Property
    End If
Next
HeaderList.List(i) = Index & ";" & vNewValue

lastError1 = Err.Description: End Property

Public Property Get SelectedItem(ByVal Index As Long) As Variant
Attribute SelectedItem.VB_MemberFlags = "400"
On Error Resume Next
If Index = 1 Then
    SelectedItem = ListView1.SelectedItem.Text
Else
    SelectedItem = ListView1.SelectedItem.ListSubItems(Index - 1).Text
End If
lastError1 = Err.Description
End Property

Public Property Let SelectedItem(ByVal Index As Long, ByVal vNewValue As Variant)
    
lastError1 = Err.Description: End Property

Public Property Get SetColumnWidth(ByVal Index As Long) As Long
Attribute SetColumnWidth.VB_MemberFlags = "400"
    On Error Resume Next
        SetColumnWidth = ListView1.ColumnHeaders(Index).Width
lastError1 = Err.Description: End Property

Public Property Let SetColumnWidth(ByVal Index As Long, ByVal vNewValue As Long)
Dim i As Long
    On Error Resume Next
        ListView1.ColumnHeaders(Index).Width = vNewValue
        For i = 0 To WidthList.ListCount - 1
            If InStr(1, WidthList.List(i), Index & ";") > 0 Then
                WidthList.List(i) = Index & ";" & vNewValue
                Exit Property
            End If
        Next
        WidthList.List(i) = Index & ";" & vNewValue
lastError1 = Err.Description:
End Property

Public Property Get ShowTools() As Boolean
    ShowTools = showTools1
lastError1 = Err.Description: End Property

Private Sub SetControl(flag As Boolean)
On Error Resume Next
If flag Then
    If IconType = LargeIcons Then
        CmdSelect.Visible = True
        CmdCancel.Visible = True
        CmdQuery.Visible = True
        CmdRefresh.Visible = True
        CmdSelect1.Visible = False
        CmdCAncel1.Visible = False
        CmdQuery1.Visible = False
        CmdRefresh1.Visible = False
    Else
        CmdSelect1.Visible = True
        CmdCAncel1.Visible = True
        CmdQuery1.Visible = True
        CmdRefresh1.Visible = True
        CmdSelect.Visible = False
        CmdCancel.Visible = False
        CmdQuery.Visible = False
        CmdRefresh.Visible = False
        'Left
        CmdSelect1.Left = 100
        CmdRefresh1.Left = CmdSelect1.Left + CmdSelect1.Width + 20
        CmdQuery1.Left = CmdRefresh1.Left + CmdRefresh1.Width + 20
        CmdCAncel1.Left = CmdQuery1.Left + CmdQuery1.Width + 20
        
        CmdSelect1.Top = 0
        CmdRefresh1.Left = 0
        CmdQuery1.Left = 0
        CmdCAncel1.Left = 0
    End If
    Call UserControl_Resize
Else
    CmdSelect.Visible = False
    CmdCancel.Visible = False
    CmdQuery.Visible = False
    'CmdOk.Visible = False
    CmdRefresh.Visible = False
    CmdSelect1.Visible = False
    CmdCAncel1.Visible = False
    CmdQuery1.Visible = False
    CmdRefresh1.Visible = False
    '-------------Set List Box-----------------------------
    If ShowText1 Then
        ListView1.Top = SearchText.Top + SearchText.Height + 50
        ListView1.Height = UserControl.ScaleHeight - SearchText.Height - 100
    Else
        ListView1.Top = 10
        ListView1.Height = UserControl.ScaleHeight
    End If
    ListView1.Width = UserControl.Width - 100
    ListView1.Left = (UserControl.Width / 2) - (ListView1.Width / 2)
End If

End Sub
Public Property Let ShowTools(ByVal vNewValue As Boolean)
    showTools1 = vNewValue
    Call UserControl_Resize
'Call SetControl(vNewValue)
lastError1 = Err.Description: End Property

Public Property Get LastError() As String
Attribute LastError.VB_MemberFlags = "400"
    LastError = lastError1
lastError1 = Err.Description: End Property

Public Property Let LastError(ByVal vNewValue As String)

lastError1 = Err.Description: End Property

Public Property Get RowCount() As Long
Attribute RowCount.VB_MemberFlags = "400"
On Error Resume Next
    RowCount = ListView1.ListItems.Count
End Property

Public Property Let RowCount(ByVal vNewValue As Long)

End Property

Public Property Get TextMatrix(ByVal Row As Long, ByVal Col As Long) As String
Attribute TextMatrix.VB_MemberFlags = "400"
    On Error Resume Next
    With ListView1
    If Col = 1 Then
        TextMatrix = .ListItems(Row).Text
    ElseIf Col > 1 Then
        TextMatrix = .ListItems(Row).SubItems(Col - 1)
    End If
    End With
End Property

Public Property Let TextMatrix(ByVal Row As Long, ByVal Col As Long, NewVal As String)
    On Error Resume Next
    With ListView1
    If Col = 1 Then
        .ListItems(Row).Text = NewVal
    ElseIf Col > 1 Then
        .ListItems(Row).SubItems(Col - 1) = NewVal
    End If
    End With
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    PropBag.WriteProperty "SHOWTOOL", showTools1, True
    PropBag.WriteProperty "ConnectionString1", ConnectionString1
    PropBag.WriteProperty "CHECKBOX", ListView1.CheckBoxes
    PropBag.WriteProperty "MULTISELECT", ListView1.MultiSelect
    PropBag.WriteProperty "MainQuery", MainQuery
    PropBag.WriteProperty "SHOWMENU", ShowMenu1, True
    PropBag.WriteProperty "FONT12", ListView1.Font
    PropBag.WriteProperty "FORECOLOR", ListView1.ForeColor
    PropBag.WriteProperty "SHOWTEXT", ShowText1, True
    PropBag.WriteProperty "ShowProgress", ShowProgress1, True
    PropBag.WriteProperty "IconType1", IconType1, 2
    PropBag.WriteProperty "ShowCaption1", True
    PropBag.WriteProperty "LBLCAPTION", Lbl_Caption.Caption
    PropBag.WriteProperty "CAPBACKCOLOR", Lbl_Caption.BackColor
    PropBag.WriteProperty "CAPFORECOLOR", Lbl_Caption.ForeColor
    PropBag.WriteProperty "BORDERCOLOR1", Border1.BorderColor
    PropBag.WriteProperty "BORDERCOLOR2", Border2.BorderColor
    PropBag.WriteProperty "CAPALIGN", Lbl_Caption.Alignment
    PropBag.WriteProperty "CAPFONT", Lbl_Caption.Font
    PropBag.WriteProperty "MovablePicklist", MovablePicklist
    PropBag.WriteProperty "LblBackground", LblBackground.BackColor
    PropBag.WriteProperty "DateFilterEndFormat1", DateFilterEndFormat1, "#"
    PropBag.WriteProperty "DateFilterStartFormat1", DateFilterStartFormat1, "#"
    PropBag.WriteProperty "DateFormat1", DateFormat1, "MM/dd/yyyy"
End Sub

Public Property Get Connected() As Boolean
Attribute Connected.VB_MemberFlags = "400"
On Error Resume Next
    If MainData.State = 1 Then
        Connected = True
    Else
        Connected = False
    End If
End Property
Public Property Let Connected(ByVal vNewValue As Boolean)
Err.Raise 78282, , "Connected property is Read only"
End Property

Public Property Get RecordCount() As Long
Attribute RecordCount.VB_MemberFlags = "400"
    RecordCount = RecordCount1
End Property

Public Property Let RecordCount(ByVal vNewValue As Long)
Err.Raise 78282, , "RecordCount property is Read only"
End Property

Public Property Get QueryString() As String
QueryString = MainQuery
End Property

Public Property Let QueryString(ByVal vNewValue As String)
MainQuery = vNewValue
PropertyChanged "QueryString"
End Property

Public Property Get ColumnCount() As Long
Attribute ColumnCount.VB_MemberFlags = "400"
    ColumnCount = CurColumnCount
End Property

Public Property Let ColumnCount(ByVal vNewValue As Long)
    Err.Raise 31231231, , "ColumnCount is ReadOnly"
End Property

Public Property Get SelectedCount() As Long
Attribute SelectedCount.VB_MemberFlags = "400"
On Error Resume Next
Dim CurSelectedCount, i As Long
With ListView1
    For i = 1 To .ListItems.Count
        If .ListItems.Item(i).Selected Then
            CurSelectedCount = CurSelectedCount + 1
        End If
    Next
End With
SelectedCount = CurSelectedCount
End Property

Public Property Let SelectedCount(ByVal vNewValue As Long)
    Err.Raise 5656, , "SelectedCount is ReadOnly"
End Property

Public Property Get MultiSelect() As Boolean
    MultiSelect = ListView1.MultiSelect
End Property

Public Property Let MultiSelect(ByVal vNewValue As Boolean)
    ListView1.MultiSelect = vNewValue
End Property

Public Property Get CheckBoxes() As Boolean
    CheckBoxes = ListView1.CheckBoxes
End Property

Public Property Let CheckBoxes(ByVal vNewValue As Boolean)
    ListView1.CheckBoxes = vNewValue
End Property

Public Property Get Selected(Index As Long) As Boolean
Attribute Selected.VB_MemberFlags = "400"
On Error GoTo errHand
    Selected = ListView1.ListItems.Item(Index).Selected
Exit Sub
errHand:
Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.Description
End Property

Public Property Let Selected(Index As Long, ByVal vNewValue As Boolean)
On Error GoTo errHand
    ListView1.ListItems.Item(Index).Selected = vNewValue
Exit Property
errHand:
Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.Description

End Property

Public Property Get ShowMenu() As Boolean
    ShowMenu = ShowMenu1
End Property

Public Property Let ShowMenu(ByVal vNewValue As Boolean)
    ShowMenu1 = vNewValue
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = ListView1.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo errHand
    ListView1.ForeColor = vNewValue
Exit Property
errHand:
Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Get Font() As IFontDisp
Set Font = ListView1.Font
End Property

Public Property Set Font(ByVal vNewValue As IFontDisp)
On Error GoTo errHand
Set ListView1.Font = vNewValue
Exit Property
errHand:
Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Get ShowText() As Boolean
    ShowText = ShowText1
End Property

Public Property Let ShowText(ByVal vNewValue As Boolean)
    ShowText1 = vNewValue
    If vNewValue Then
        SearchText.Visible = True
    Else
        SearchText.Visible = False
    End If
    Call UserControl_Resize
End Property

Public Property Get ShowProgress() As Boolean
ShowProgress = ShowProgress1
End Property

Public Property Let ShowProgress(ByVal vNewValue As Boolean)
ShowProgress1 = vNewValue
End Property

Public Property Get IconType() As IconSize
    IconType = IconType1
End Property

Public Property Let IconType(ByVal vNewValue As IconSize)
    IconType1 = vNewValue
    PropertyChanged "IconType"
    Call UserControl_Resize
End Property

Public Property Get ShowCaption() As Boolean
    ShowCaption = ShowCaption1
End Property

Public Property Let ShowCaption(ByVal vNewValue As Boolean)
    ShowCaption1 = vNewValue
    Lbl_Caption.Visible = vNewValue
    LblBackground.Visible = vNewValue
    Border1.Visible = vNewValue
    Border2.Visible = vNewValue
    Call Resize
End Property

Public Property Get PicklistCaption() As String
    PicklistCaption = Lbl_Caption
End Property

Public Property Let PicklistCaption(ByVal vNewValue As String)
    Lbl_Caption = vNewValue
End Property

Public Property Get CaptionBackcolor() As OLE_COLOR
    CaptionBackcolor = Lbl_Caption.BackColor
End Property

Public Property Let CaptionBackcolor(ByVal vNewValue As OLE_COLOR)
    Lbl_Caption.BackColor = vNewValue
End Property

Public Property Get CaptionForeColor() As OLE_COLOR
    CaptionForeColor = Lbl_Caption.ForeColor
End Property

Public Property Let CaptionForeColor(ByVal vNewValue As OLE_COLOR)
    Lbl_Caption.ForeColor = vNewValue
End Property

Public Property Get Border1Color() As OLE_COLOR
    Border1Color = Border1.BorderColor
End Property

Public Property Let Border1Color(ByVal vNewValue As OLE_COLOR)
    Border1.BorderColor = vNewValue
End Property

Public Property Get Border2Color() As OLE_COLOR
    Border1Color = Border1.BorderColor
End Property

Public Property Let Border2Color(ByVal vNewValue As OLE_COLOR)
    Border2.BorderColor = vNewValue
End Property

Public Property Get CaptionAlignment() As AlignmentConstants
    CaptionAlignment = Lbl_Caption.Alignment
End Property

Public Property Let CaptionAlignment(ByVal vNewValue As AlignmentConstants)
    Lbl_Caption.Alignment = vNewValue
End Property

Public Property Get CaptionFont() As StdFont
    Set CaptionFont = Lbl_Caption.Font
End Property

Public Property Set CaptionFont(ByVal vNewValue As StdFont)
    Set Lbl_Caption.Font = vNewValue
End Property

Public Property Get Movable() As Boolean
    Movable = MovablePicklist
End Property

Public Property Let Movable(ByVal vNewValue As Boolean)
    MovablePicklist = vNewValue
End Property

Public Property Get CaptionBackgroundColor() As OLE_COLOR
    CaptionBackgroundColor = LblBackground.BackColor
End Property

Public Property Let CaptionBackgroundColor(ByVal vNewValue As OLE_COLOR)
    LblBackground.BackColor = vNewValue
End Property

Public Property Get EnableButton(Index As Long) As Boolean
    Select Case Index
        Case 1
            EnableButton = CmdSelect.Enabled
        Case 2
            EnableButton = CmdRefresh.Enabled
        Case 3
            EnableButton = CmdQuery.Enabled
        Case 4
            EnableButton = CmdCancel.Enabled
    End Select
End Property

Public Property Let EnableButton(Index As Long, ByVal vNewValue As Boolean)
    Select Case Index
        Case 1
            CmdSelect.Enabled = vNewValue
            CmdSelect1.Enabled = vNewValue
        Case 2
            CmdRefresh.Enabled = vNewValue
            CmdRefresh1.Enabled = vNewValue
        Case 3
            CmdQuery.Enabled = vNewValue
            CmdQuery1.Enabled = vNewValue
        Case 4
            CmdCancel.Enabled = vNewValue
            CmdCAncel1.Enabled = vNewValue
    End Select
End Property

Public Property Get DateFilterStartFormat() As String
    DateFilterStartFormat = DateFilterStartFormat1
End Property

Public Property Let DateFilterStartFormat(ByVal vNewValue As String)
If Trim$(vNewValue) = "" Then Exit Property
    DateFilterStartFormat1 = vNewValue
End Property
Public Property Get DateFilterEndFormat() As String
    DateFilterEndFormat = DateFilterEndFormat1
End Property

Public Property Let DateFilterEndFormat(ByVal vNewValue As String)
If Trim$(vNewValue) = "" Then Exit Property
    DateFilterEndFormat1 = vNewValue
End Property
Public Property Get DateFormat() As String
    DateFormat = DateFormat1
End Property

Public Property Let DateFormat(ByVal vNewValue As String)
If Trim$(vNewValue) = "" Then Exit Property
    DateFormat1 = vNewValue
End Property

