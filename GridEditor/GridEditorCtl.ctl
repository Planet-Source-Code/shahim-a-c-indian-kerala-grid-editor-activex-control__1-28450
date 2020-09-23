VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.UserControl GridEditor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11625
   ControlContainer=   -1  'True
   PropertyPages   =   "GridEditorCtl.ctx":0000
   ScaleHeight     =   3525
   ScaleWidth      =   11625
   ToolboxBitmap   =   "GridEditorCtl.ctx":0011
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "Print"
      Height          =   360
      Left            =   15
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton CmdFull 
      Caption         =   "Full View"
      Height          =   360
      Left            =   825
      TabIndex        =   19
      Top             =   30
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton Cmd_Close 
      Caption         =   "Close"
      Height          =   360
      Left            =   1635
      TabIndex        =   18
      Top             =   30
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   4200
      ScaleHeight     =   720
      ScaleWidth      =   2670
      TabIndex        =   16
      Top             =   1260
      Visible         =   0   'False
      Width           =   2670
      Begin MSForms.Label Label2 
         Height          =   285
         Left            =   15
         TabIndex        =   17
         Top             =   225
         Width           =   2610
         Caption         =   "Printing, Please wait..."
         Size            =   "4604;503"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.CheckBox chkColWidth 
      Caption         =   "Resize Col &widths to fill page"
      Height          =   195
      Left            =   2820
      TabIndex        =   15
      Top             =   105
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox picScroll 
      Height          =   3375
      Left            =   6570
      ScaleHeight     =   3315
      ScaleWidth      =   4635
      TabIndex        =   11
      Top             =   1470
      Visible         =   0   'False
      Width           =   4695
      Begin VB.HScrollBar hscScroll 
         Height          =   255
         LargeChange     =   15
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3000
         Width           =   4575
      End
      Begin VB.VScrollBar vscScroll 
         Height          =   2535
         LargeChange     =   15
         Left            =   4035
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   75
         Width           =   255
      End
      Begin VB.PictureBox picTarget 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   15
         ScaleHeight     =   2625
         ScaleWidth      =   3825
         TabIndex        =   14
         Top             =   -15
         Width           =   3855
      End
   End
   Begin VB.ListBox NoValues 
      Height          =   840
      Left            =   7695
      TabIndex        =   10
      Top             =   2310
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   540
      Top             =   0
   End
   Begin VB.ListBox MaxL 
      Height          =   840
      Left            =   10215
      TabIndex        =   6
      Top             =   1365
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.ListBox OnlyNo 
      Height          =   840
      Left            =   10320
      TabIndex        =   5
      Top             =   2340
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.ListBox ObjectList 
      Height          =   840
      Left            =   8880
      TabIndex        =   4
      Top             =   2415
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox EnableList 
      Height          =   840
      Left            =   8880
      TabIndex        =   3
      Top             =   1395
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TxtEditor 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1245
      TabIndex        =   2
      Top             =   2580
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox Pict1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1110
      ScaleHeight     =   495
      ScaleWidth      =   1695
      TabIndex        =   1
      Top             =   1845
      Visible         =   0   'False
      Width           =   1695
      Begin MSComCtl2.DTPicker DatePicker 
         Height          =   270
         Left            =   -285
         TabIndex        =   8
         Top             =   165
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   476
         _Version        =   393216
         Format          =   62783489
         CurrentDate     =   37100
      End
      Begin MSComCtl2.DTPicker TimePicker 
         Height          =   270
         Left            =   60
         TabIndex        =   9
         Top             =   255
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   476
         _Version        =   393216
         Format          =   62783490
         CurrentDate     =   37100
      End
      Begin MSForms.ComboBox CmbEditor 
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   135
         Width           =   1470
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2593;450"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         BorderColor     =   -2147483643
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSFlexGridLib.MSFlexGrid EditorGrid 
      Height          =   3015
      Left            =   105
      TabIndex        =   0
      Top             =   720
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      RowHeightMin    =   285
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GridEditorCtl.ctx":0323
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GridEditorCtl.ctx":057D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu rem1 
      Caption         =   "Remove"
      Visible         =   0   'False
      Begin VB.Menu remr 
         Caption         =   "Remove Row"
      End
   End
End
Attribute VB_Name = "GridEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'************************************************************************
'*  Programmer   :   SHAHIM.A.C                                         *
'*  Control      :   GridEditor Control                                 *
'*  Purpose      :   User can directly edit the msflexgrid              *
'************************************************************************
Option Explicit
Enum ObjectType
    TextBox = 1
    ComboBox = 2
    DateBox = 3
    TimeBox = 4
    CheckBox = 5
End Enum
Enum PrinterOrientation
    LandScape = 1
    Portrait = 2
End Enum
Private Type RetrieveNum
    FirstVal As Long
    secondval As Variant
    Thirdval As Integer
    LastVal As Double
End Type
'Public dateBox1 As DTPicker
Private RetN As RetrieveNum
'Public EditBox As New Editor
Private ShowEdit As Boolean
Private Advance As Boolean
Private RowVal, ColVal As Long
Dim temp, Temp2, PrevValue As String
Private Cancelled As Boolean
Private Menutoshow As Boolean
Private CreateNewRows1 As Boolean
Private AddingItem As Boolean
Private ChangingCols As Boolean
Private CheckBoxisthere As Boolean
Private NoofPages1 As Long
Private EditorColor1 As OLE_COLOR, EditorForeColor1 As OLE_COLOR
Private CellEditArray() As String
Private RowEditArray() As String
Private CheckEnableEdit As Boolean 'Used in then Insertcontrol sub
Private PrintOrient As PrinterOrientation

'---------------------Events for MsFlexGrid
Public Event Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Public Event AdvanceToNextCell(Cancel As Boolean)
Public Event ExitFocus(Cancel As Boolean)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click()
Public Event DblClick()
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event RowColChange()
Public Event Scroll()
Public Event SelChange()
Public Event LeaveCell()
Public Event EnterCell()
Public Event EditorValidate(Cancel As Boolean)
Public Event NewRow(Cancel As Boolean)
Attribute NewRow.VB_MemberFlags = "40"
Public Event BeforeNewRow(Cancel As Boolean)
Public Event AfterNewRowCreated()
Public Event NewPage(LastPrintedRow As Long, CurrentPage As Long)
'Public Event DragOver(Source As Control, x As Single, y As Single, State As Integer)
'Public Event DragDrop(Source As Control, x As Single, y As Single)

'---------------------Events for Datepicker
Public Event DateBoxCloseUp()
Public Event DateBoxDropDown()
Public Event Format(CallbackField As String, FormattedString As String)
Public Event FormatSize(CallbackField As String, Size As Integer)
Public Event ComboDropButtonClick()
Public Event TimeCallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
Public Event EditorChange()
Public Event DatePickerCallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

Private WithEvents cTP As clsTablePrint 'For Print Grid
Attribute cTP.VB_VarHelpID = -1
'The dimensions of the DIN A4 paper size in Twips:
Const A4Height = 16840, A4Width = 11907

'To get the scroll width:
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CYHSCROLL = 3
Private Const SM_CXVSCROLL = 2

Private Sub InsertCheckBox(Optional Value As Integer)
On Error Resume Next

    With EditorGrid
    If .Row = .FixedRows - 1 Then Exit Sub
        If Value = 0 Then
            .CellPictureAlignment = 4
            Set .CellPicture = ImageList1.ListImages(1).Picture
            If .CellBackColor = 0 Then .CellBackColor = vbWhite
            .CellforeColor = .CellBackColor
            .TextMatrix(.Row, .Col) = 0
        Else
            .CellPictureAlignment = 4
            Set .CellPicture = ImageList1.ListImages(2).Picture
            If .CellBackColor = 0 Then .CellBackColor = vbWhite
            .CellforeColor = .CellBackColor
            .TextMatrix(.Row, .Col) = 1
        End If
    End With
End Sub
Private Sub RetNos(Colno As Long)
    RetN.FirstVal = 0: RetN.secondval = "": RetN.Thirdval = 0: RetN.LastVal = 0
    NoValues.Clear
    Dim i, j, k, Col1 As Long, temp, temp1 As String
    With OnlyNo
        For i = 0 To .ListCount - 1
            For j = 1 To Len(.List(i))
                temp = temp & Mid(.List(i), j, 1)
                    If Mid(.List(i), j, 1) = ";" Then
                        If Val(Left(temp, Len(temp) - 1)) = Colno Then
                            NoValues.AddItem Colno
                            For k = j + 1 To Len(.List(i))
                                temp1 = temp1 & Mid(.List(i), k, 1)
                                If Mid(.List(i), k, 1) = ";" Then
                                    NoValues.AddItem Left(temp1, Len(temp1) - 1)
                                    temp1 = ""
                                End If
                            Next
                        Else
                            GoTo skip
                        End If
                    End If
                    'temp = ""
            Next
skip:
            temp = "": temp1 = ""
        Next
End With
With NoValues
    RetN.FirstVal = Val(.List(0))
    RetN.secondval = .List(1)
    RetN.Thirdval = Val(.List(2))
    RetN.LastVal = Val(.List(3))
End With
    
End Sub

Public Sub ShowAbout()
Attribute ShowAbout.VB_UserMemId = -552
    Frm_About_Editor.Show vbModal
End Sub
Public Sub AddItem(Item As String)
Attribute AddItem.VB_Description = "To add an item into the grid"
Dim CurRow, i, CurCol As Long

On Error Resume Next
    With EditorGrid
    .AddItem Item
    If Not CheckBoxisthere Then Exit Sub
    AddingItem = True
        CurRow = .Row: CurCol = .Col
        .Row = .Rows - 1
            For i = .FixedCols To .Cols - 1
                .Col = i
                If CheckWhichCtl = 5 Then
                    If Val(.Text) = 0 Then
                    InsertCheckBox 0
                Else
                    InsertCheckBox 1
                End If
                End If
            Next
        .Row = CurRow: .Col = CurCol
        .RowHeight(EditorGrid.Rows - 1) = .RowHeightMin 'EditorGrid.RowHeight(EditorGrid.Row - 1)
    Call HideEditor
    End With
    AddingItem = False
Exit Sub
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Sub

Public Property Get EditorFont() As StdFont
On Error Resume Next
    Set EditorFont = CmbEditor.Font
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
    
End Property

Public Property Set EditorFont(ByVal vNewValue As StdFont)
On Error Resume Next
    Set CmbEditor.Font = vNewValue
    Set TxtEditor.Font = vNewValue
    PropertyChanged "EditorFont"
    Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property
Public Property Get GridFont() As StdFont
On Error Resume Next
    Set GridFont = EditorGrid.Font
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear

End Property

Public Property Set GridFont(ByVal vNewValue As StdFont)
On Error Resume Next
    Set EditorGrid.Font = vNewValue
    PropertyChanged "GridFont"
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Private Sub RetrieveVal(List As VB.ListBox, Col As Long)
'This procedure is used to get the information from a particular listbox
RetN.FirstVal = 0
RetN.secondval = ""
    Dim i, j, Col1 As Long
    With List
        For i = 0 To .ListCount - 1
            For j = 1 To Len(.List(i))
                If Mid(.List(i), j, 1) = ";" Then
                    Col1 = Left(.List(i), j - 1)
                    If Col = Col1 Then
                        RetN.FirstVal = Col1 'Left(.List(i), j - 1)
                        RetN.secondval = Right(.List(i), Len(.List(i)) - j)
                        Exit Sub
                    End If
                End If
            Next
        Next
    End With
End Sub
Public Sub RemoveItem(Row As Long)
On Error Resume Next
    EditorGrid.RemoveItem Row
    HideEditor
Exit Sub
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Sub
Private Function CheckforCombo() As Integer
'This function is to check which object to be place in a given cell
    Dim i, j, flag, Col As Long
    
    With ObjectList
        For i = 0 To .ListCount - 1
        Col = 0
        flag = 0
            For j = 1 To Len(.List(i))
                If Mid(.List(i), j, 1) = ";" Then
                    Col = Left(.List(i), j - 1)
                    flag = Right(.List(i), Len(.List(i)) - j)
                    If Col = EditorGrid.Col And flag = 1 Then
                        CheckforCombo = 1
                        Exit Function
                    End If
                End If
            Next
        Next
    End With
End Function

Private Function CheckWhichCtl() As Integer
'This function is to check which object to be place in a given cell
    Dim i, j, Col As Long
    Dim flag As Single
    With ObjectList
        For i = 0 To .ListCount - 1
        Col = 0
        flag = 0
            For j = 1 To Len(.List(i))
                If Mid(.List(i), j, 1) = ";" Then
                    Col = Left(.List(i), j - 1)
                    flag = Right(.List(i), Len(.List(i)) - j)
                    If Col = EditorGrid.Col Then
                        CheckWhichCtl = flag
'                    If flag = 1 Then
'                        CheckWhichCtl = 1
'                    ElseIf flag = 2 Then
'                        CheckWhichCtl = 2
'                    Else: CheckWhichCtl = 0
'                    End If
                        Exit Function
                    End If
                End If
            Next
        Next
    End With
End Function

Public Property Get TextMatrix(ByVal Row As Long, ByVal Col As Long) As String
Attribute TextMatrix.VB_MemberFlags = "400"
'To get the Text of a particular Cell
    TextMatrix = EditorGrid.TextMatrix(Row, Col)
End Property

Public Property Let TextMatrix(ByVal Row As Long, ByVal Col As Long, ByVal vNewValue As String)
On Error Resume Next
'To let the Text of a particular Cell
Dim CurRow, CurCol As Long
    With EditorGrid
        HideEditor
        .TextMatrix(Row, Col) = vNewValue
        If Not CheckBoxisthere Then GoTo skip
        CurRow = .Row: CurCol = .Col
        .Row = Row: .Col = Col
            If CheckWhichCtl = 5 Then
                If vNewValue = 0 Then
                    InsertCheckBox 0
                Else
                    InsertCheckBox 1
                End If
            End If
        .Row = CurRow: .Col = CurCol
skip:
        If Not CheckEnabled Or Not ShowEdit Then Exit Property
        InsertControl
    End With
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get SetColObject(ByVal Col As Long) As ActifOcx.ObjectType '  ObjectType
Attribute SetColObject.VB_Description = "To set the object to be displayed in a  particular cell as a combo or a text\r\n"
Attribute SetColObject.VB_MemberFlags = "400"
'To set the object to be displayed in a  particular cell

    
End Property

Public Property Let SetColObject(ByVal Col As Long, ByVal vNewValue As ActifOcx.ObjectType)
'To set the object to be displayed in a  particular cell
'If Col > EditorGrid.Cols - 1 Then
'    Err.Raise 101, , "Invalid Column" & vbCr & "Column " & Col & " does not exist"
'    Exit Property
'End If
Dim Col1 As Long
    Dim i, j As Long
    With ObjectList
        For i = 0 To .ListCount - 1
            For j = 1 To Len(.List(i))
            If Mid(.List(i), j, 1) = ";" Then
                Col1 = Left(.List(i), j - 1)
                If Col = Col1 Then
                    .RemoveItem i
                End If
            End If
            Next
        Next
        ObjectList.AddItem Col & ";" & vNewValue
        If Val(Right(vNewValue, 1)) = 5 Then
            CheckBoxisthere = True
        End If
    End With
End Property
Public Property Get EnableEditing(ByVal Col As Long) As Boolean
If Col > EditorGrid.Cols - 1 Then
    Err.Raise 423423, , "Column does not exists"
End If
Dim i, j As Long
With EnableList
    For i = 0 To .ListCount - 1
        For j = 1 To Len(.List(i))
        If Mid(.List(i), j, 1) = ";" Then
            If Col = Val(Left(.List(i), j - 1)) Then
                EnableEditing = CBool(Right(.List(i), Len(.List(i)) - j))
                Exit Property
            End If
        End If
        Next
    Next
EnableEditing = True
End With
End Property
Public Property Let EnableEditing(ByVal Col As Long, ByVal vNewValue As Boolean)
Attribute EnableEditing.VB_Description = "Enable or Disable the control to be placed in the grid"
Attribute EnableEditing.VB_MemberFlags = "400"
'Enable or Disable the control to be placed in the grid
'If Col > EditorGrid.Cols - 1 Then
 '   Err.Raise 102, , "Invalid column number." & vbCr & "Column " & Col & " does not exist."
'End If
Dim Col1 As Long
    Dim i, j As Long
    With EnableList
        For i = 0 To .ListCount - 1
            For j = 1 To Len(.List(i))
            If Mid(.List(i), j, 1) = ";" Then
                Col1 = Left(.List(i), j - 1)
                If Col = Col1 Then
                    .RemoveItem i
                End If
            End If
            Next
        Next
End With
    EnableList.AddItem Col & ";" & vNewValue
End Property

Private Sub chkColWidth_Click()
Call PrintView
End Sub

Private Sub CmbEditor_Change()
On Error Resume Next
EditorGrid = CmbEditor
With EditorGrid
    If .Row = .Rows - 1 And CreateNewRows1 And Not AddingItem Then 'And Not Trim$(.TextMatrix(.Row, .FixedCols)) = "" Then
        Dim Cancel As Boolean
        RaiseEvent BeforeNewRow(Cancel)
        If Cancel Then GoTo skip
        .AddItem ""
        .RowHeight(.Row + 1) = .RowHeightMin ' .RowHeight(.Row)
        RaiseEvent AfterNewRowCreated
    End If
End With
skip:
RaiseEvent EditorChange
End Sub

Private Sub CmbEditor_Click()
    RaiseEvent Click
End Sub

Private Sub CmbEditor_DblClick(Cancel As MSForms.ReturnBoolean)
    RaiseEvent DblClick
End Sub

Private Sub CmbEditor_DropButtonClick()
    RaiseEvent ComboDropButtonClick
End Sub


Private Sub CmbEditor_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim k As Integer
Dim Cancel As Boolean
With EditorGrid
    Select Case KeyCode
        Case 13
            Call IncrementCell
            If Not ShowEdit Or Not CheckEnabled Then
                HideEditor
            Else
                If TxtEditor.Visible Then
                    TxtEditor.SetFocus
                ElseIf Pict1.Visible Then
                    If CmbEditor.Visible Then
                        CmbEditor.SetFocus
                    ElseIf DatePicker.Visible Then
                        DatePicker.SetFocus
                    Else: TimePicker.SetFocus
                    End If
                End If
            End If
        Case vbKeyEscape
            CmbEditor = PrevValue
        Case vbKeyUp
            If Shift = 2 Then
                RaiseEvent AdvanceToNextCell(Cancel)
                If Cancel Then Exit Sub
                If Not .Row = .FixedRows Then
                    temp = .TextMatrix(.Row - 1, .Col)
                    .Row = .Row - 1
                    Timer1.Enabled = True
                End If
            End If
        Case vbKeyDown
            If Shift = 2 Then
                RaiseEvent AdvanceToNextCell(Cancel)
                If Cancel Then Exit Sub
                If Not .Row = .Rows - 1 Then
                    temp = .TextMatrix(.Row + 1, .Col)
                    .Row = .Row + 1
                    Timer1.Enabled = True
                End If
            End If
        Case vbKeyLeft
            RaiseEvent AdvanceToNextCell(Cancel)
            If Cancel Then Exit Sub
            If .Col > .FixedCols And CmbEditor.SelStart = 0 Then
                .Col = .Col - 1
                Exit Sub
            End If
'            If .Col = .FixedCols And .Row > .FixedRows And CmbEditor.SelStart = 0 Then
'                .Row = .Row - 1
'                CmbEditor.SetFocus
'            End If
        Case vbKeyRight
            RaiseEvent AdvanceToNextCell(Cancel)
            If Cancel Then Exit Sub
            If .Col < .Cols - 1 And CmbEditor.SelStart = Len(CmbEditor) Then
                .Col = .Col + 1
                Exit Sub
            End If
'            If .Col = .Cols - 1 And .Row < .Rows - 1 And CmbEditor.SelStart = Len(CmbEditor) Then
'                .Row = .Row + 1
'                CmbEditor.SetFocus
'            End If
        Case Else
'            Call RetNos(.Col)
'            If RetN.FirstVal = .Col And RetN.secondval = True Then
'                KeyCode = OnlyNumbers(CmbEditor, KeyCode, RetN.Thirdval, RetN.LastVal)
'            End If
    End Select
End With
k = KeyCode
RaiseEvent KeyDown(k, Shift)
End Sub

Private Sub CmbEditor_KeyPress(KeyAscii As MSForms.ReturnInteger)
NoValues.Clear
    Call RetNos(EditorGrid.Col)
    If RetN.FirstVal = EditorGrid.Col And RetN.secondval = True Then
        KeyAscii = OnlyNumbers(CmbEditor, KeyAscii, RetN.Thirdval, RetN.LastVal)
    End If
    NoValues.Clear
    RaiseEvent KeyPress(Val(KeyAscii))
End Sub



Private Sub CmbEditor_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    RaiseEvent KeyUp(Int(KeyCode), Shift)
End Sub

Private Sub CmbEditor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmbEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub CmbEditor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Cmd_Close_Click()
Cmd_Close.Visible = False
Cmd_Print.Visible = False
EditorGrid.Visible = True
picScroll.Visible = False
chkColWidth.Visible = False
CmdFull.Visible = False
End Sub

Private Sub Cmd_Print_Click()
Call PrintGrid
End Sub

Private Sub CmdFull_Click()
On Error Resume Next
Screen.MousePointer = vbHourglass
    
    Call InitializePictureBoxToPreview
    Set cTP = New clsTablePrint
    Call cmdRefresh_ClicktoPreview
    Frm_Grid_Preview.Show
Screen.MousePointer = vbDefault

End Sub

Private Sub cTP_NewPage(objOutput As Object, TopMarginAlreadySet As Boolean, bCancel As Boolean, ByVal lLastPrintedRow As Long)
NoofPages1 = NoofPages1 + 1
RaiseEvent NewPage(lLastPrintedRow, NoofPages1)
    'The class wants a new page, look what to do
    If TypeOf objOutput Is Printer Then
        Printer.NewPage
    Else 'We are printing on the PictureBox !
        objOutput.CurrentY = objOutput.ScaleHeight
        'Simply increase the height of the PicBox here
        ' (very simple, but looks bad in "real" applications)
        objOutput.Height = objOutput.Height + A4Height
        'Draw a line to show the new page:
        objOutput.Line (0, objOutput.CurrentY)-(objOutput.ScaleWidth, objOutput.CurrentY), &H808080
        
        'Set the CurrentY to the position the class should continie with drawing and...
        objOutput.CurrentY = objOutput.CurrentY + cTP.MarginTop
        '... tell it to do so:
        TopMarginAlreadySet = True
        
        'Set the ScrollBar Max properties:
        SetScrollBars
    End If
End Sub

Private Sub DatePicker_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    RaiseEvent DatePickerCallbackKeyDown(KeyCode, Shift, CallbackField, CallbackDate)
End Sub

Private Sub DatePicker_Change()
On Error Resume Next
If Not IsNull(DatePicker) Then EditorGrid = DatePicker
With EditorGrid
    If .Row = .Rows - 1 And CreateNewRows1 And Not AddingItem Then   'And Not Trim$(.TextMatrix(.Row, .FixedCols)) = "" Then
    Dim Cancel As Boolean
        RaiseEvent BeforeNewRow(Cancel)
        If Cancel Then GoTo skip
        .AddItem ""
        .RowHeight(.Row + 1) = .RowHeightMin
        RaiseEvent AfterNewRowCreated
        '.RowHeight(.Row + 1) = .RowHeight(.Row)
    End If
End With
skip:
RaiseEvent EditorChange
Exit Sub
errHand:
Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Private Sub DatePicker_Click()
    RaiseEvent Click
End Sub

Private Sub DatePicker_CloseUp()
    RaiseEvent DateBoxCloseUp
End Sub


Private Sub DatePicker_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub DatePicker_DropDown()
    RaiseEvent DateBoxDropDown
End Sub


Private Sub DatePicker_Format(ByVal CallbackField As String, FormattedString As String)
    RaiseEvent Format(CallbackField, FormattedString)
End Sub


Private Sub DatePicker_FormatSize(ByVal CallbackField As String, Size As Integer)
    RaiseEvent FormatSize(CallbackField, Size)
End Sub


Private Sub DatePicker_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Cancel As Boolean
With EditorGrid
    Select Case KeyCode
        Case 13
            Call IncrementCell
            If Not ShowEdit Or Not CheckEnabled Then
                HideEditor
            Else
                If TxtEditor.Visible Then
                    TxtEditor.SetFocus
                ElseIf Pict1.Visible Then
                    If CmbEditor.Visible Then
                        CmbEditor.SetFocus
                    ElseIf DatePicker.Visible Then
                        DatePicker.SetFocus
                    Else: TimePicker.SetFocus
                    End If
                End If
            End If
        Case vbKeyEscape
        On Error Resume Next
            DatePicker = DateValue(Format(PrevValue, "dd/mm/yyyy"))
'        Case vbKeyUp
'            RaiseEvent AdvanceToNextCell(Cancel)
'            If Cancel Then Exit Sub
'            If Not .Row = .FixedRows Then
'                Temp2 = .TextMatrix(.Row - 1, .Col)
'                .Row = .Row - 1
'                Timer2.Enabled = True
'            End If
'        Case vbKeyDown
'            RaiseEvent AdvanceToNextCell(Cancel)
'            If Cancel Then Exit Sub
'            If Not .Row = .Rows - 1 Then
'                Temp2 = .TextMatrix(.Row + 1, .Col)
'                .Row = .Row + 1
'                Timer2.Enabled = True
'            End If
'        Case vbKeyLeft
'            RaiseEvent AdvanceToNextCell(Cancel)
'            If Cancel Then Exit Sub
'            If .Col > .FixedCols And CmbEditor.SelStart = 0 Then
'                .Col = .Col - 1
'                Exit Sub
'            End If
''            If .Col = .FixedCols And .Row > .FixedRows And CmbEditor.SelStart = 0 Then
''                .Row = .Row - 1
''                CmbEditor.SetFocus
''            End If
'        Case vbKeyRight
'            RaiseEvent AdvanceToNextCell(Cancel)
'            If Cancel Then Exit Sub
'            If .Col < .Cols - 1 And CmbEditor.SelStart = Len(CmbEditor) Then
'                .Col = .Col + 1
'                Exit Sub
'            End If
''            If .Col = .Cols - 1 And .Row < .Rows - 1 And CmbEditor.SelStart = Len(CmbEditor) Then
''                .Row = .Row + 1
''                CmbEditor.SetFocus
''            End If
'        Case Else
'            Call RetrieveVal(OnlyNo, .Col)
'            If RetN.FirstVal = .Col And RetN.secondval = True Then
'                KeyCode = OnlyNumbers(TxtEditor, KeyCode)
'            End If
    End Select
End With
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub DatePicker_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub DatePicker_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub DatePicker_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub DatePicker_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub DatePicker_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub EditorGrid_Click()
With EditorGrid
If Not CheckBoxisthere Then GoTo skip
If CheckWhichCtl = 5 And CheckEnabled Then
    If Val(.Text) = 0 Then
        InsertCheckBox 1
    Else: InsertCheckBox 0
    End If
End If
End With
skip:
    InsertControl
    RaiseEvent Click
End Sub

Private Sub EditorGrid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    RaiseEvent Compare(Row1, Row2, Cmp)
End Sub

Private Sub EditorGrid_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub EditorGrid_EnterCell()
If Cancelled Then
    EditorGrid.Row = RowVal
    EditorGrid.Col = ColVal
    InsertControl
    Cancelled = False
    GoTo skip
End If
If Not ShowEdit Or Not CheckEnabled Then
    HideEditor
    GoTo skip
End If

'Dim Cancel As Boolean
'RaiseEvent AdvanceToNextCell(Cancel)
'If Cancel Then
'    Cancel = False
'    EditorGrid.Row = RowVal
'    EditorGrid.Col = ColVal
'    InsertControl
'Exit Sub
'End If
    InsertControl
skip:
    RaiseEvent EnterCell
End Sub

Private Sub EditorGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub EditorGrid_KeyPress(KeyAscii As Integer)
Dim keyasc As Integer
keyasc = KeyAscii
Select Case KeyAscii
'    Case vbKeyBack
'    Case vbKeyDelete
'    Case vbKeySpace
'    Case vbKeyF1 To vbKeyF12
'        KeyAscii = 0
    Case vbKeyScrollLock
        KeyAscii = 0
    Case vbKeyEscape
        KeyAscii = 0
    Case 13
        Call IncrementCell
        If Not ShowEdit Or Not CheckEnabled Then
            Call HideEditor
            Exit Sub
        Else
            If TxtEditor.Visible Then
                    TxtEditor.SetFocus
                ElseIf Pict1.Visible Then
                    If CmbEditor.Visible Then
                        CmbEditor.SetFocus
                    ElseIf DatePicker.Visible Then
                        DatePicker.SetFocus
                    Else: TimePicker.SetFocus
                    End If
                End If
        End If
    'Case vbKeySpace
        
    Case Else
            With EditorGrid
            If Not CheckBoxisthere Then GoTo skip
                If CheckWhichCtl = 5 And CheckEnabled Then
                    KeyAscii = 0
                    Call EditorGrid_Click
                    GoTo skip
                End If
                If Not CheckEnabled Or Not ShowEdit Then Exit Sub
                Call RetNos(.Col)
                If RetN.FirstVal = EditorGrid.Col And RetN.secondval = True Then
                    KeyAscii = OnlyNumbers(EditorGrid, KeyAscii, RetN.Thirdval, RetN.LastVal)
                End If
                    EditorGrid = Chr(KeyAscii)
                Call InsertControl
            End With
    End Select
skip:
InsertControl
RaiseEvent KeyPress(keyasc)
End Sub

Private Sub EditorGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub EditorGrid_LeaveCell()
If Cancelled Then Exit Sub
Debug.Print "Row : " & EditorGrid.Row
Debug.Print "Col : " & EditorGrid.Col
Dim Cancel As Boolean
RaiseEvent AdvanceToNextCell(Cancel)
If Cancel Then
    Cancelled = True
    EditorGrid.Row = RowVal
    EditorGrid.Col = ColVal
End If
RaiseEvent LeaveCell
'MsgBox "hi"
End Sub

Private Sub EditorGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub EditorGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub EditorGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub EditorGrid_RowColChange()
    RaiseEvent RowColChange
End Sub

Private Sub EditorGrid_Scroll()
    RaiseEvent Scroll
    HideEditor
End Sub

Private Sub test_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub EditorGrid_SelChange()
    RaiseEvent SelChange
End Sub

Private Sub hscScroll_Change()
picTarget.Left = -hscScroll.Value * 120
End Sub

Private Sub TimePicker_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    RaiseEvent TimeCallbackKeyDown(KeyCode, Shift, CallbackField, CallbackDate)
End Sub

Private Sub TimePicker_Change()
On Error Resume Next
If Not IsNull(TimePicker) Then EditorGrid = Format(TimePicker, "HH:MM AM/PM")
With EditorGrid
    If .Row = .Rows - 1 And CreateNewRows1 And Not AddingItem Then 'And Not Trim$(.TextMatrix(.Row, .FixedCols)) = "" Then
        Dim Cancel As Boolean
        RaiseEvent BeforeNewRow(Cancel)
        If Cancel Then GoTo skip
        .AddItem ""
        .RowHeight(.Row + 1) = .RowHeightMin
        RaiseEvent AfterNewRowCreated
        '.RowHeight(.Row + 1) = .RowHeight(.Row)
    End If
End With
skip:
RaiseEvent EditorChange
End Sub


Private Sub TimePicker_Click()
    RaiseEvent Click
End Sub

Private Sub TimePicker_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub TimePicker_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Cancel As Boolean
With EditorGrid
    Select Case KeyCode
        Case 13
            Call IncrementCell
            If Not ShowEdit Or Not CheckEnabled Then
                HideEditor
            Else
                If TxtEditor.Visible Then
                    TxtEditor.SetFocus
                ElseIf Pict1.Visible Then
                    If CmbEditor.Visible Then
                        CmbEditor.SetFocus
                    ElseIf DatePicker.Visible Then
                        DatePicker.SetFocus
                    Else: TimePicker.SetFocus
                    End If
                End If
            End If
        Case vbKeyEscape
        On Error Resume Next
            TimePicker = TimeValue(Format(PrevValue, "hh:mm:ss"))
'        Case vbKeyUp
'            RaiseEvent AdvanceToNextCell(Cancel)
'            If Cancel Then Exit Sub
'            If Not .Row = .FixedRows Then
'                Temp2 = .TextMatrix(.Row - 1, .Col)
'                .Row = .Row - 1
'                Timer2.Enabled = True
'            End If
'        Case vbKeyDown
'            RaiseEvent AdvanceToNextCell(Cancel)
'            If Cancel Then Exit Sub
'            If Not .Row = .Rows - 1 Then
'                Temp2 = .TextMatrix(.Row + 1, .Col)
'                .Row = .Row + 1
'                Timer2.Enabled = True
'            End If
'        Case vbKeyLeft
'            RaiseEvent AdvanceToNextCell(Cancel)
'            If Cancel Then Exit Sub
'            If .Col > .FixedCols And CmbEditor.SelStart = 0 Then
'                .Col = .Col - 1
'                Exit Sub
'            End If
''            If .Col = .FixedCols And .Row > .FixedRows And CmbEditor.SelStart = 0 Then
''                .Row = .Row - 1
''                CmbEditor.SetFocus
''            End If
'        Case vbKeyRight
'            RaiseEvent AdvanceToNextCell(Cancel)
'            If Cancel Then Exit Sub
'            If .Col < .Cols - 1 And CmbEditor.SelStart = Len(CmbEditor) Then
'                .Col = .Col + 1
'                Exit Sub
'            End If
''            If .Col = .Cols - 1 And .Row < .Rows - 1 And CmbEditor.SelStart = Len(CmbEditor) Then
''                .Row = .Row + 1
''                CmbEditor.SetFocus
''            End If
'        Case Else
'            Call RetrieveVal(OnlyNo, .Col)
'            If RetN.FirstVal = .Col And RetN.secondval = True Then
'                KeyCode = OnlyNumbers(TxtEditor, KeyCode)
'            End If
    End Select
End With
RaiseEvent KeyDown(KeyCode, Shift)
End Sub


Private Sub TimePicker_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub TimePicker_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub TimePicker_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub TimePicker_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub TimePicker_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()
CmbEditor = temp
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
DatePicker = DateValue(Format(Temp2, DatePicker.Format))
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()

End Sub

Private Sub TxtEditor_Change()
On Error Resume Next
EditorGrid = Trim(TxtEditor)
With EditorGrid
    If .Row = .Rows - 1 And CreateNewRows1 And Not AddingItem Then  'And Not Trim$(.TextMatrix(.Row, .FixedCols)) = "" Then
        Dim Cancel As Boolean
        RaiseEvent BeforeNewRow(Cancel)
        If Cancel Then GoTo skip
        .AddItem ""
        .RowHeight(.Row + 1) = .RowHeightMin
        RaiseEvent AfterNewRowCreated
        '.RowHeight(.Row + 1) = .RowHeight(.Row)
    End If
End With
skip:
NoValues.Clear
    Call RetNos(EditorGrid.Col)
    If RetN.FirstVal = EditorGrid.Col And RetN.secondval = True Then
    EditorGrid = Val(TxtEditor)
        If Val(TxtEditor) > Val(RetN.LastVal) Or IsNumeric(TxtEditor) = False Then
            TxtEditor = ""
        End If
    End If
NoValues.Clear
RaiseEvent EditorChange
Exit Sub
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'Private Sub TxtEditor_KeyDown(KeyCode As Integer, Shift As Integer)
''Select Case KeyCode
''    Case 13
''        Call IncrementCell
''        If Not ShowEdit Or Not CheckEnabled Then
''            HideEditor
''        Else
''            If TxtEditor.Visible Then
''                TxtEditor.SetFocus
''            ElseIf Pict1.Visible Then
''                CmbEditor.SetFocus
''            End If
''        End If
''
''End Select
'End Sub



Public Property Get Rows() As Long
Attribute Rows.VB_Description = "To get or set the rows of a grid"
    Rows = EditorGrid.Rows
End Property

Public Property Let Rows(ByVal vNewValue As Long)
On Error Resume Next
Dim i, j, CurRow, CurCol, NoOfRows As Long
    With EditorGrid
        NoOfRows = .Rows
        .Rows = vNewValue
        Call HideEditor
If Not CheckBoxisthere Then Exit Property
EditorGrid.Visible = False
        CurRow = .Row: CurCol = .Col
            For i = NoOfRows To .Rows - 1
                '.RowHeight(i) =  .RowHeight(i) * 1.2
                .Row = i
                    For j = .FixedCols To .Cols - 1
                        .Col = j
                        If CheckWhichCtl = 5 Then
                            InsertCheckBox
                        End If
                    Next
                Next
EditorGrid.Visible = False
        .Row = CurRow: .Col = CurCol
    End With
    Call HideEditor
End Property

Public Property Get Cols() As Long
    Cols = EditorGrid.Cols
End Property

Public Property Let Cols(ByVal vNewValue As Long)
On Error Resume Next
Dim i, j, CurRow, CurCol, NoOfCols As Long
    With EditorGrid
        NoOfCols = .Cols
        .Cols = vNewValue
        Call HideEditor
        If Not CheckBoxisthere Then Exit Property
        EditorGrid.Visible = False
        ChangingCols = True
        CurRow = .Row: CurCol = .Col
            For i = .FixedRows To .Rows - 1
                .Row = i
                For j = NoOfCols To .Cols - 1
                    .Col = j
                    If CheckWhichCtl = 5 Then
                        InsertCheckBox
                    End If
                Next
            Next
        .Row = CurRow: .Col = CurCol
        EditorGrid.Visible = True
    End With
ChangingCols = False
Exit Property
errHand:
MsgBox Err.Description, vbExclamation
End Property

Public Sub HideEditor(Optional flag As Boolean = True)
If flag Then
    TxtEditor.Visible = False
    Pict1.Visible = False
    DatePicker.Visible = False
    TimePicker.Visible = False
    CmbEditor.Visible = False
Else
    TxtEditor.Visible = True
    Pict1.Visible = True
End If
End Sub
Private Sub focusSet(ob As TextBox)
On Error Resume Next
    With ob
        .SelStart = 0
        .SelLength = Len(ob)
        .SetFocus
    End With
End Sub
Private Sub InsertTextBox()
On Error Resume Next
    HideEditor
    If ChangingCols Then Exit Sub
    With EditorGrid
        TxtEditor.Width = .CellWidth - 10
        TxtEditor.Height = .CellHeight
        TxtEditor.Top = .CellTop + .Top ' - 10
        TxtEditor.Left = .Left + .CellLeft
        PrevValue = EditorGrid
        TxtEditor = .Text
        TxtEditor.SelStart = Len(TxtEditor)
        RowVal = .Row
        ColVal = .Col
    End With
    TxtEditor.Visible = True
    TxtEditor.SetFocus
    RetrieveVal MaxL, EditorGrid.Col
        If RetN.FirstVal = EditorGrid.Col Then
            TxtEditor.MaxLength = RetN.secondval
        Else
            TxtEditor.MaxLength = 0
        End If
End Sub
Private Sub InsertCombo()
On Error Resume Next
    HideEditor
    With EditorGrid
        Pict1.Top = .CellTop + .Top ' - 10
        Pict1.Left = .Left + .CellLeft
        Pict1.Width = .CellWidth
        Pict1.Height = .CellHeight
        RowVal = .Row
        ColVal = .Col
    End With
    
    With CmbEditor
        .Top = 0
        .Left = 0
        .Height = Pict1.Height
        .Width = Pict1.Width
        PrevValue = EditorGrid
        '.ListIndex = 0
        .Text = EditorGrid
        .SelStart = Len(CmbEditor)
        RetrieveVal MaxL, EditorGrid.Col
        If RetN.FirstVal = EditorGrid.Col Then
            .MaxLength = RetN.secondval
        Else
            .MaxLength = 0
        End If
    End With
    Pict1.Visible = True
    CmbEditor.Visible = True
    CmbEditor.SetFocus
End Sub
Private Sub InsertDateCombo()
On Error Resume Next
    HideEditor
    With EditorGrid
        Pict1.Top = .CellTop + .Top ' - 10
        Pict1.Left = .Left + .CellLeft
        Pict1.Width = .CellWidth
        Pict1.Height = .CellHeight
        RowVal = .Row
        ColVal = .Col
    End With
    
    With DatePicker
        .Top = 0
        .Left = 0
        .Height = Pict1.Height
        .Width = Pict1.Width
        PrevValue = EditorGrid
        .Value = DateValue(Format(EditorGrid, "dd/mm/yyyy"))
    End With
    Pict1.Visible = True
    DatePicker.Visible = True
    DatePicker.SetFocus
End Sub
Private Sub InsertTimeCombo()
On Error Resume Next
    HideEditor
    With EditorGrid
        Pict1.Top = .CellTop + .Top ' - 10
        Pict1.Left = .Left + .CellLeft
        Pict1.Width = .CellWidth
        Pict1.Height = .CellHeight
        RowVal = .Row
        ColVal = .Col
    End With
    
    With TimePicker
        .Top = 0
        .Left = 0
        .Height = Pict1.Height
        .Width = Pict1.Width
        PrevValue = EditorGrid
        .Value = TimeValue(Format(EditorGrid, "hh:mm:ss"))
    End With
    Pict1.Visible = True
    TimePicker.Visible = True
    TimePicker.SetFocus
End Sub
Public Function CheckEnabled() As Boolean
'This function is to check whether for a particular cell editing is enabled or not.
    CheckEnabled = True
    Dim i, j, Col As Long
    Dim flag As String
    With EnableList
        For i = 0 To .ListCount - 1
        Col = 0
        flag = True
            For j = 1 To Len(.List(i))
                If Mid(.List(i), j, 1) = ";" Then
                    Col = Left(.List(i), j - 1)
                    flag = Right(.List(i), Len(.List(i)) - j)
                    If Col = EditorGrid.Col And flag = "False" Then
                        CheckEnabled = False
                        Exit Function
                    End If
                End If
            Next
        Next
    If EnableCellEdit(EditorGrid.Row, EditorGrid.Col) = True Then CheckEnabled = True
    End With
    
End Function
Private Sub InsertControl()
On Error Resume Next
    With EditorGrid
        If Not CheckEnabled Or Not ShowEdit Or EditorGrid.Rows = 0 Or .Row = 0 _
        Or Not EnableCellEdit(.Row, .Col) Or Not EnableRowEdit(.Row) Then 'or .Row<=.FixedCols-1
'Change according Enmedit val property.If true then show else hidden
        CheckEnableEdit = True  'To check to forcefully enable a cell if the cell is already disabled
        If EnableCellEdit(.Row, .Col) Then GoTo skip
            Call HideEditor
            EditorGrid.SetFocus
            Exit Sub
        End If
skip:
        'If ObType.Col = .Col And ObType.colObject = ComboBox Then
        Dim Tmp As Single
        Tmp = CheckWhichCtl
        If Tmp = 2 Then
            Call InsertCombo
        ElseIf Tmp = 3 Then
            Call InsertDateCombo
        ElseIf Tmp = 4 Then
            Call InsertTimeCombo
        ElseIf Tmp = 5 Then
            Call InsertTextBox
            HideEditor
        Else
            Call InsertTextBox
        End If
    End With
    
End Sub
Private Sub IncrementCell()
On Error Resume Next
Dim Cancel As Boolean
RaiseEvent AdvanceToNextCell(Cancel)
If Cancel = True Then Exit Sub
'Call CheckEnabled
    With EditorGrid
    If .Row < .FixedRows Then Exit Sub
        If .Col >= .FixedCols And .Col < .Cols - 1 Then 'ie., when the columns > fixedcolumns and <=no of cols
            .Col = .Col + 1
            If Not CheckEnabled Then Call IncrementCell
            .SetFocus
            Exit Sub
        End If
        
        If .Col = .Cols - 1 And .Row < .Rows - 1 Then
            .Row = .Row + 1
            .Col = .FixedCols
            If Not CheckEnabled Then Call IncrementCell
            .SetFocus
            Exit Sub
        End If
        If Not CreateNewRows1 Then Exit Sub
        If .Col = .Cols - 1 And .Row = .Rows - 1 And ShowEdit Then
        RaiseEvent BeforeNewRow(Cancel)
        If Cancel Then Exit Sub
            .AddItem ""
        RaiseEvent AfterNewRowCreated
            .RowHeight(.Row + 1) = .RowHeightMin
            '.RowHeight(.Row + 1) = .RowHeight(.Row)
            .Row = .Row + 1
            .Col = .FixedCols
            .SetFocus
            If Not CheckEnabled Then Call IncrementCell
            Exit Sub
        End If
        
    End With
End Sub

Private Sub IncrementCell1()
'Call CheckEnabled
    With EditorGrid
    If .Row = 0 Then Exit Sub
        If .Col >= .FixedCols And .Col < .Cols - 1 Then 'ie., when the columns > fixedcolumns and <=no of cols
            .Col = .Col + 1
            .SetFocus
            Exit Sub
        End If
        
        If .Col = .Cols - 1 And .Row < .Rows - 1 Then
            .Row = .Row + 1
            .Col = .FixedCols
            .SetFocus
            Exit Sub
        End If
        
        If .Col = .Cols - 1 And .Row = .Rows - 1 And ShowEdit Then
        Dim Cancel As Boolean
        RaiseEvent BeforeNewRow(Cancel)
        If Cancel Then Exit Sub
            .AddItem ""
        RaiseEvent AfterNewRowCreated
            .RowHeight(.Row + 1) = .RowHeightMin
            '.RowHeight(.Row + 1) = .RowHeight(.Row)
            .Row = .Row + 1
            .Col = .FixedCols
            .SetFocus
            Exit Sub
        End If
        
    End With
End Sub

Public Property Get ShowEditor() As Boolean
    ShowEditor = ShowEdit
End Property

Public Property Let ShowEditor(ByVal vNewValue As Boolean)
        ShowEdit = vNewValue
        If Not vNewValue Then Call HideEditor
        PropertyChanged "ShowEditor"
End Property

Private Sub TxtEditor_Click()
    RaiseEvent Click
End Sub

Private Sub TxtEditor_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub TxtEditor_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Cancel As Boolean
With EditorGrid
    Select Case KeyCode
        Case 13
            Call IncrementCell
            If Not ShowEdit Or Not CheckEnabled Then
                HideEditor
            Else
                If TxtEditor.Visible Then
                    TxtEditor.SetFocus
                ElseIf Pict1.Visible Then
                    If CmbEditor.Visible Then
                        CmbEditor.SetFocus
                    ElseIf DatePicker.Visible Then
                        DatePicker.SetFocus
                    Else: TimePicker.SetFocus
                    End If
                End If
            End If
        Case vbKeyEscape
            TxtEditor = PrevValue
        Case vbKeyUp
            RaiseEvent AdvanceToNextCell(Cancel)
            If Cancel Then Exit Sub
                If Not .Row = .FixedRows Then
                    .Row = .Row - 1
                End If
            Case vbKeyDown
            RaiseEvent AdvanceToNextCell(Cancel)
            If Cancel Then Exit Sub
                If Not .Row = .Rows - 1 Then
                    .Row = .Row + 1
                End If
            Case vbKeyLeft
            RaiseEvent AdvanceToNextCell(Cancel)
            If Cancel Then Exit Sub
                If .Col > .FixedCols And TxtEditor.SelStart = 0 Then
                    .Col = .Col - 1
                End If
            Case vbKeyRight
            RaiseEvent AdvanceToNextCell(Cancel)
            If Cancel Then Exit Sub
                If .Col < .Cols - 1 And TxtEditor.SelStart = Len(TxtEditor) Then
                    .Col = .Col + 1
                End If
        Case Else
'            Call RetNos(.Col)
'            If RetN.FirstVal = .Col And RetN.secondval = True Then
'                KeyCode = OnlyNumbers(TxtEditor, KeyCode, RetN.Thirdval, RetN.LastVal)
'            End If
    End Select
End With
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Function OnlyNumbers(Ctl As Object, ByVal KeyAscii As Integer, Optional ByVal Value As Double = 2, Optional ByVal Limit As Double = 99999.99) As Integer
On Error Resume Next
Dim i As Long
Dim leftNo, RightNo As String
Dim tempno, Tmp As String
tempno = ""
    'Function for Only numbers to be entered
'    If Val(Value) > 0 Then
'MsgBox Left(Ctl, Ctl.SelStart) & vbCr & Right(Ctl, Len(Ctl) - Ctl.SelStart)
'temp = Left(Ctl, Ctl.SelStart) & Asc(Chr(KeyAscii)) & Right(Ctl, Len(Ctl) - Ctl.SelStart)
'MsgBox temp
        
        'If Val(Ctl) > Limit Then OnlyNumbers = 0: Ctl = "": Exit Function
'    ElseIf Val(Ctl) > 99999.99 Then OnlyNumbers = 0: Exit Function
'    End If
    OnlyNumbers = KeyAscii
    Select Case KeyAscii
        Case vbKeyBack, 13, vbKeyDelete
        If KeyAscii = vbKeyBack Then
            Tmp = Left(Ctl, Ctl.SelStart - 1) & vbCr & Right(Ctl, Len(Ctl) - Ctl.SelStart)
            If Val(Tmp) > Limit Then
                OnlyNumbers = 0: Exit Function
            End If
        End If
        Case 48 To 57
            If Val(Ctl & Chr(KeyAscii)) > Limit Then
                OnlyNumbers = 0
                Exit Function
            End If
        If Val(Ctl) > 0 Then
            leftNo = Left(Trim$(Ctl), Ctl.SelStart)
            RightNo = Right(Trim$(Ctl), Len(Ctl) - Ctl.SelStart)
            If Val(leftNo & Chr(KeyAscii) & RightNo) > Limit Then
                OnlyNumbers = 0
                Exit Function
            End If
        End If
            For i = 1 To Len(Ctl)
                If Mid$(Ctl, i, 1) = "." Then
                    If Len(Left(Trim$(Ctl), Len(Trim$(Ctl)) - i)) = Value Then
                        If Ctl.SelStart >= i Then
                            OnlyNumbers = 0
                        End If
                    End If
                End If
            Next i
        
        Case Asc%(".")
            If Value = 0 Then OnlyNumbers = 0: Exit Function
            For i = 1 To Len(Ctl)
                If Mid$(Ctl, i, 1) = "." Then
                    OnlyNumbers = 0
                End If
            Next i
        Case Else
            OnlyNumbers = 0
    End Select
End Function

Private Sub TxtEditor_KeyPress(KeyAscii As Integer)
    NoValues.Clear
    Call RetNos(EditorGrid.Col)
    If RetN.FirstVal = EditorGrid.Col And RetN.secondval = True Then
        KeyAscii = OnlyNumbers(TxtEditor, KeyAscii, RetN.Thirdval, RetN.LastVal)
    End If
    NoValues.Clear
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub TxtEditor_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub TxtEditor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub TxtEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub TxtEditor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub TxtEditor_Validate(Cancel As Boolean)
    RaiseEvent EditorValidate(Cancel)
End Sub

Private Sub UserControl_EnterFocus()

On Error Resume Next
With EditorGrid
    If .Rows = 0 Or .FixedRows = .Rows Then Exit Sub
    If Not CheckEnabled Or Not ShowEdit Or .Rows = 0 Then
        HideEditor
        Exit Sub
    Else
    InsertControl
    End If
End With
End Sub

Private Sub UserControl_ExitFocus()
On Error Resume Next
Dim Cancel As Boolean
RaiseEvent ExitFocus(Cancel)
If Cancel Then UserControl.SetFocus
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
    Advance = True
    CreateNewRows1 = True
    NoofPages1 = 1
    ReDim CellEditArray(0 To 1)
    ReDim RowEditArray(0 To 1)
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    With PropBag
        EditorGrid.Rows = .ReadProperty("ROWS")
        EditorGrid.Cols = .ReadProperty("COLS", 4)
        ShowEditor = .ReadProperty("SHOWEDITOR", True)
        Set CmbEditor.Font = .ReadProperty("FONT1") ', EditorGrid.Font
        Set EditorGrid.Font = .ReadProperty("GRIDFONT")
'        EditorGrid.Appearance = .ReadProperty("APPEAR")
        EditorGrid.BackColor = .ReadProperty("BACKCOLOR")
        EditorGrid.BackColorBkg = .ReadProperty("BACKCOLORBKG")
        EditorGrid.BackColorFixed = .ReadProperty("BACKCOLORFIXED")
        EditorGrid.BackColorSel = .ReadProperty("BACKCOLORSEL", EditorGrid.BackColorSel)
        'EditorGrid.BorderStyle = .ReadProperty("BORDERSTYLE")
        EditorGrid.FillStyle = .ReadProperty("FILLSTYLE")
        EditorGrid.FixedCols = .ReadProperty("FIXEDCOLS")
        EditorGrid.FixedRows = .ReadProperty("FIXEDROWS")
        EditorGrid.FocusRect = .ReadProperty("FOCUSRECT")
        EditorGrid.ForeColor = .ReadProperty("FORECOLOR")
        EditorGrid.ForeColorFixed = .ReadProperty("FORECOLORFIXED")
        EditorGrid.ForeColorSel = .ReadProperty("FORECOLORSEL")
        EditorGrid.FormatString = .ReadProperty("FORMATSTRING")
        EditorGrid.GridColor = .ReadProperty("GRIDCOLOR")
        EditorGrid.GridColorFixed = .ReadProperty("GRIDCOLORFIXED")
        EditorGrid.GridLines = .ReadProperty("GRIDLINES")
        EditorGrid.GridLinesFixed = .ReadProperty("GRIDLINESFIXED")
        EditorGrid.GridLineWidth = .ReadProperty("GRIDLINEWIDTH")
        EditorGrid.MergeCells = .ReadProperty("MERGECELLS")
        Set EditorGrid.MouseIcon = .ReadProperty("MOUSEICON")
        CmbEditor.AutoWordSelect = .ReadProperty("ComboAutoWordSelect")
        EditorGrid.MousePointer = .ReadProperty("MOUSEPOINTER")
        EditorGrid.PictureType = .ReadProperty("PICTURETYPE")
        EditorGrid.Redraw = .ReadProperty("REDRAW")
        EditorGrid.RightToLeft = .ReadProperty("RIGHTTOLEFT")
        EditorGrid.ScrollBars = .ReadProperty("SCROLLBARS")
        EditorGrid.ScrollTrack = .ReadProperty("SCROLLTRACK")
        EditorGrid.SelectionMode = .ReadProperty("SELECTIONMODE")
        EditorGrid.TextStyle = .ReadProperty("TEXTSTYLE")
        EditorGrid.TextStyleFixed = .ReadProperty("TEXTSTYLEFIXED")
        EditorGrid.ToolTipText = .ReadProperty("TOOLTIPTEXT")
        EditorGrid.RowHeightMin = .ReadProperty("ROWHEIGHTMIN", EditorGrid.RowHeightMin)
        CreateNewRows1 = .ReadProperty("CREATENEWROWS")
        CmbEditor.MatchEntry = .ReadProperty("ComboMachEntry", CmbEditor.Style)
        CmbEditor.Style = .ReadProperty("ComboStyle")
        DatePicker.CalendarBackColor = .ReadProperty("CALENDERBACKCOLOR", DatePicker.CalendarBackColor)
        DatePicker.CalendarForeColor = .ReadProperty("CALENDERFORECOLOR", DatePicker.CalendarForeColor)
        DatePicker.CalendarTitleBackColor = .ReadProperty("CALENDERTITLEBACKCOLOR", DatePicker.CalendarTitleBackColor)
        DatePicker.CalendarTitleForeColor = .ReadProperty("CALENDERTITLEFORECLOR", DatePicker.CalendarTitleForeColor)
        DatePicker.CalendarTrailingForeColor = .ReadProperty("CALENDERTRAILING", DatePicker.CalendarTrailingForeColor)
        DatePicker.Format = .ReadProperty("CALENDERFORMAT", DatePicker.Format)
        DatePicker.CustomFormat = .ReadProperty("CALNDERCUSTOMFORMAT", DatePicker.CustomFormat)
        DatePicker.UpDown = .ReadProperty("CALENDERUPDOWN", DatePicker.UpDown)
        EditorGrid.WordWrap = .ReadProperty("WORDWRAP", EditorGrid.WordWrap)
        CmbEditor.MatchRequired = .ReadProperty("COMBOMATCHREQUIRED", CmbEditor.MatchRequired)
        CmbEditor.ShowDropButtonWhen = .ReadProperty("SHOWDROPBUTTONWHEN", CmbEditor.ShowDropButtonWhen)
        EditorGrid.AllowUserResizing = .ReadProperty("ALLOWUSERRESIZING", EditorGrid.AllowUserResizing)
        EditorGrid.Enabled = .ReadProperty("ENABLED", EditorGrid.Enabled)
        EditorColor1 = .ReadProperty("SetEditorColor", vbWhite)
        EditorForeColor1 = .ReadProperty("EditorForeColor1", vbBlack)
        PrintOrient = .ReadProperty("PrintOrient", 1)
        'TimePicker.CustomFormat = .ReadProperty("TIMEFORMAT", TimePicker.CustomFormat)
'        DatePicker.Value = .ReadProperty("DATETIME")
'        DatePicker.MinDate = .ReadProperty("DATEMIN")
'        DatePicker.MaxDate = .ReadProperty("DATEMAX")
'        DatePicker.Format = .ReadProperty("DATEFORMAT")
'        DatePicker.CustomFormat = .ReadProperty("CUSTOMFORMAT")
'        DatePicker.CheckBox = .ReadProperty("CHECKBOX")
'        DatePicker.UpDown = .ReadProperty("UPDOWN")


'        TimePick.Value = .ReadProperty("TIMEPICK")
'         Menutoshow = .ReadProperty("MNUSHOW")
    End With
    
    Dim i As Long
    With EditorGrid
    For i = 0 To .Rows - 1
        .RowHeight(i) = .RowHeightMin ' .RowHeight(i) * 1.2
    Next
    End With
    If EditorColor1 <= 0 Then EditorColor1 = vbWhite
    
    TxtEditor.BackColor = EditorColor1
    CmbEditor.BackColor = EditorColor1
    TxtEditor.ForeColor = EditorForeColor1
    CmbEditor.ForeColor = EditorForeColor1
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
With EditorGrid
    .Left = 0
    .Top = 0
    .Height = UserControl.ScaleHeight
    .Width = UserControl.ScaleWidth
End With
On Error Resume Next
With picScroll
    .Width = ScaleWidth
    .Top = Cmd_Print.Height + 70
    .Left = 0
    .Height = ScaleHeight - (Cmd_Print.Height + 30)
End With
    
    hscScroll.Max = (picTarget.Width - picScroll.ScaleWidth + vscScroll.Width) / 120 + 1
    vscScroll.Max = (picTarget.Height - picScroll.ScaleHeight + hscScroll.Height) / 120 + 1
   
    hscScroll.Left = 0
    hscScroll.Width = UserControl.ScaleWidth - (vscScroll.Width + 40)
    hscScroll.Top = UserControl.ScaleHeight - (hscScroll.Height + picScroll.Top + 40)
    
    vscScroll.Top = 0
    vscScroll.Height = UserControl.ScaleHeight - (hscScroll.Height + Cmd_Print.Height + 40)
    vscScroll.Left = UserControl.ScaleWidth - vscScroll.Width - 50
    
End Sub

Private Sub UserControl_Show()
Dim i, j, CurRow, CurCol As Long
With EditorGrid
CurRow = .Row: CurCol = .Col
If GetSetting("GridEditor", "Validate", "Show") = "0" Then
    Frm_About_Editor.Show vbModal
End If
If Not CheckBoxisthere Then Exit Sub
For i = .FixedRows To .Rows - 1
    .Row = i
        For j = .FixedCols To .Cols - 1
            .Col = j
            If CheckWhichCtl = 5 Then
                InsertCheckBox
            End If
        Next
    Next
    .Row = CurRow: .Col = CurCol
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
With PropBag
    .WriteProperty "ROWS", EditorGrid.Rows, 2
    .WriteProperty "COLS", EditorGrid.Cols, 2
    .WriteProperty "SHOWEDITOR", ShowEditor, True
    .WriteProperty "FONT1", TxtEditor.Font
    .WriteProperty "GRIDFONT", EditorGrid.Font
'    .WriteProperty "APPEAR", EditorGrid.Appearance, 1
    .WriteProperty "BACKCOLOR", EditorGrid.BackColor
    .WriteProperty "BACKCOLORBKG", EditorGrid.BackColorBkg
    .WriteProperty "BACKCOLORFIXED", EditorGrid.BackColorFixed
    .WriteProperty "BACKCOLORSEL", EditorGrid.BackColorSel
    '.WriteProperty "BORDERSTYLE", EditorGrid.BorderStyle
    .WriteProperty "FILLSTYLE", EditorGrid.FillStyle
    .WriteProperty "FIXEDCOLS", EditorGrid.FixedCols
    .WriteProperty "FIXEDROWS", EditorGrid.FixedRows
    .WriteProperty "FOCUSRECT", EditorGrid.FocusRect
    .WriteProperty "FORECOLOR", EditorGrid.ForeColor
    .WriteProperty "FORECOLORFIXED", EditorGrid.ForeColorFixed
    .WriteProperty "FORECOLORSEL", EditorGrid.ForeColorSel
    .WriteProperty "FORMATSTRING", EditorGrid.FormatString
    .WriteProperty "GRIDCOLOR", EditorGrid.GridColor
    .WriteProperty "GRIDCOLORFIXED", EditorGrid.GridColorFixed
    .WriteProperty "GRIDLINES", EditorGrid.GridLines
    .WriteProperty "GRIDLINESFIXED", EditorGrid.GridLinesFixed
    .WriteProperty "GRIDLINEWIDTH", EditorGrid.GridLineWidth
    .WriteProperty "MERGECELLS", EditorGrid.MergeCells
    .WriteProperty "MOUSEICON", EditorGrid.MouseIcon
    .WriteProperty "MOUSEPOINTER", EditorGrid.MousePointer
    .WriteProperty "PICTURETYPE", EditorGrid.PictureType
    .WriteProperty "REDRAW", EditorGrid.Redraw
    .WriteProperty "RIGHTTOLEFT", EditorGrid.RightToLeft
    .WriteProperty "ROWHEIGHTMIN", EditorGrid.RowHeightMin
    .WriteProperty "SCROLLBARS", EditorGrid.ScrollBars
    .WriteProperty "SCROLLTRACK", EditorGrid.ScrollTrack
    .WriteProperty "SELECTIONMODE", EditorGrid.SelectionMode
    .WriteProperty "TEXTSTYLE", EditorGrid.TextStyle
    .WriteProperty "TEXTSTYLEFIXED", EditorGrid.TextStyleFixed
    .WriteProperty "TOOLTIPTEXT", EditorGrid.ToolTipText
    .WriteProperty "WORDWRAP", EditorGrid.WordWrap
    .WriteProperty "CREATENEWROWS", CreateNewRows1
    .WriteProperty "ComboAutoWordSelect", CmbEditor.AutoWordSelect
    .WriteProperty "ComboMachEntry", CmbEditor.MatchEntry
    .WriteProperty "ComboStyle", CmbEditor.Style
    .WriteProperty "CALENDERBACKCOLOR", DatePicker.CalendarBackColor
    .WriteProperty "CALENDERFORECOLOR", DatePicker.CalendarForeColor
    .WriteProperty "CALENDERTITLEBACKCOLOR", DatePicker.CalendarTitleBackColor
    .WriteProperty "CALENDERTITLEFORECLOR", DatePicker.CalendarTitleForeColor
    .WriteProperty "CALENDERTRAILING", DatePicker.CalendarTrailingForeColor
    .WriteProperty "CALENDERFORMAT", DatePicker.Format
    .WriteProperty "CALNDERCUSTOMFORMAT", DatePicker.CustomFormat
    .WriteProperty "CALENDERUPDOWN", DatePicker.UpDown
    .WriteProperty "COMBOMATCHREQUIRED", CmbEditor.MatchRequired
    .WriteProperty "SHOWDROPBUTTONWHEN", CmbEditor.ShowDropButtonWhen
    .WriteProperty "ALLOWUSERRESIZING", EditorGrid.AllowUserResizing
    .WriteProperty "ENABLED", EditorGrid.Enabled
    .WriteProperty "SetEditorColor", EditorColor1, vbWhite
    .WriteProperty "EditorForeColor1", EditorForeColor1, vbBlack
    .WriteProperty "PrintOrient", PrintOrient, 1
    '.WriteProperty "TIMEFORMAT", TimePicker.CustomFormat
    
'    .WriteProperty "DATETIME", DatePicker.Value
'    .WriteProperty "DATEMIN", DatePicker.MinDate
'    .WriteProperty "DATEMAX", DatePicker.MaxDate
'    .WriteProperty "DATEFORMAT", DatePicker.Format
'    .WriteProperty "CUSTOMFORMAT", DatePicker.CustomFormat
'    .WriteProperty "CHECKBOX", DatePicker.CheckBox
'    .WriteProperty "UPDOWN", DatePicker.UpDown
'    .WriteProperty "TIMEPICK", TimePick.Value
End With
End Sub

Public Property Get Row() As Long
Attribute Row.VB_MemberFlags = "400"
    Row = EditorGrid.Row
End Property

Public Property Let Row(ByVal vNewValue As Long)
On Error Resume Next
With EditorGrid
If .Row > .Rows - 1 Then
    Err.Raise 103, "Row value out of range"
End If
    .Row = vNewValue
End With
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property
Public Property Get Col() As Long
Attribute Col.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Col.VB_MemberFlags = "400"
    Col = EditorGrid.Col
End Property

Public Property Let Col(ByVal vNewValue As Long)
On Error Resume Next
With EditorGrid
'If .Col > .Cols - 1 Then
'    Err.Raise 103, "Column value out of range"
'End If
    .Col = vNewValue
End With
Call HideEditor
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

'Public Property Get Appearance() As AppearanceSettings
'    Appearance = EditorGrid.Appearance
'End Property
'
'Public Property Let Appearance(ByVal vNewValue As AppearanceSettings)
'    EditorGrid.Appearance = vNewValue
'End Property

Public Property Get BackColor() As OLE_COLOR
On Error Resume Next
     BackColor = EditorGrid.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
On Error Resume Next
    EditorGrid.BackColor = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get BackColorBkg() As OLE_COLOR
'    BackColorBkg = EditorGrid.BackColorBkg
End Property

Public Property Let BackColorBkg(ByVal vNewValue As OLE_COLOR)
On Error Resume Next
    EditorGrid.BackColorBkg = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get BackColorFixed() As OLE_COLOR
On Error Resume Next
    BackColorFixed = EditorGrid.BackColorFixed
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Let BackColorFixed(ByVal vNewValue As OLE_COLOR)
On Error Resume Next
    EditorGrid.BackColorFixed = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get BackColorSel() As OLE_COLOR
    'BackColorSel = EditorGrid.BackColorSel
End Property

Public Property Let BackColorSel(ByVal vNewValue As OLE_COLOR)
On Error Resume Next
    EditorGrid.BackColorSel = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property
'
'Public Property Get BorderStyle() As BorderStyleSettings
'    BorderStyle = EditorGrid.BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal vNewValue As BorderStyleSettings)
'    EditorGrid.BorderStyle = vNewValue
'End Property

Public Property Get FillStyle() As FillStyleSettings
'    FillStyle = EditorGrid.FillStyle
End Property

Public Property Let FillStyle(ByVal vNewValue As FillStyleSettings)
On Error Resume Next
    EditorGrid.FillStyle = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get FixedCols() As Long
On Error Resume Next
    FixedCols = EditorGrid.FixedCols
End Property

Public Property Let FixedCols(ByVal vNewValue As Long)
On Error Resume Next
    EditorGrid.FixedCols = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get FixedRows() As Long
'Set EditorGrid = EditorGrid
    FixedRows = EditorGrid.FixedRows
End Property

Public Property Let FixedRows(ByVal vNewValue As Long)
On Error Resume Next
    EditorGrid.FixedRows = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get FocusRect() As FocusRectSettings
    FocusRect = EditorGrid.FocusRect
End Property

Public Property Let FocusRect(ByVal vNewValue As FocusRectSettings)
On Error Resume Next
    EditorGrid.FocusRect = vNewValue
    Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
    ForeColorFixed = EditorGrid.ForeColorFixed
End Property

Public Property Let ForeColorFixedFixed(ByVal vNewValue As OLE_COLOR)
On Error Resume Next
    EditorGrid.ForeColorFixed = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get ForeColorSel() As OLE_COLOR
    ForeColorSel = EditorGrid.ForeColorSel
End Property

Public Property Let ForeColorSel(ByVal vNewValue As OLE_COLOR)
On Error Resume Next
    EditorGrid.ForeColorSel = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Set or get the text"
Attribute Text.VB_MemberFlags = "400"
    Text = EditorGrid.Text
End Property

Public Property Let Text(ByVal vNewValue As String)
On Error Resume Next
    EditorGrid.Text = vNewValue
    Call HideEditor
    If Not CheckBoxisthere Then Exit Property
    If CheckWhichCtl = 5 Then
        If Val(vNewValue) = 0 Then
            InsertCheckBox 0
        Else
            InsertCheckBox 1
        End If
    End If
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get FormatString() As String
    FormatString = EditorGrid.FormatString
End Property

Public Property Let FormatString(ByVal vNewValue As String)
On Error Resume Next
    EditorGrid.FormatString = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get GridColor() As OLE_COLOR
    GridColor = EditorGrid.GridColor
End Property

Public Property Let GridColor(ByVal vNewValue As OLE_COLOR)
On Error Resume Next
    EditorGrid.GridColor = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get GridColorFixed() As OLE_COLOR
Attribute GridColorFixed.VB_MemberFlags = "40"
    GridColorFixed = EditorGrid.GridColorFixed
End Property

Public Property Let GridColorFixed(ByVal vNewValue As Variant)
On Error Resume Next
    EditorGrid.GridColorFixed = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property
Public Property Get GridColourFixed() As OLE_COLOR
    GridColorFixed = EditorGrid.GridColorFixed
End Property

Public Property Let GridColourFixed(ByVal vNewValue As OLE_COLOR)
On Error Resume Next
    EditorGrid.GridColorFixed = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property
Public Property Get GridLines() As GridLineSettings
    GridLines = EditorGrid.GridLines
End Property

Public Property Let GridLines(ByVal vNewValue As GridLineSettings)
On Error Resume Next
    EditorGrid.GridLines = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get GridLinesFixed() As GridLineSettings
    GridLinesFixed = EditorGrid.GridLinesFixed
End Property

Public Property Let GridLinesFixed(ByVal vNewValue As GridLineSettings)
On Error Resume Next
    EditorGrid.GridLinesFixed = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get GridLineWidth() As Integer
    GridLineWidth = EditorGrid.GridLineWidth
End Property

Public Property Let GridLineWidth(ByVal vNewValue As Integer)
On Error Resume Next
    EditorGrid.GridLineWidth = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get Highlight() As HighLightSettings
    Highlight = EditorGrid.Highlight
End Property

Public Property Let Highlight(ByVal vNewValue As HighLightSettings)
On Error Resume Next
    EditorGrid.Highlight = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get MergeCells() As MergeCellsSettings
    MergeCells = EditorGrid.MergeCells
End Property

Public Property Let MergeCells(ByVal vNewValue As MergeCellsSettings)
On Error Resume Next
    EditorGrid.MergeCells = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get PictureType() As PictureTypeSettings
Attribute PictureType.VB_MemberFlags = "400"
    PictureType = EditorGrid.PictureType
End Property

Public Property Let PictureType(ByVal vNewValue As PictureTypeSettings)
On Error Resume Next
    EditorGrid.PictureType = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get Redraw() As Boolean
    Redraw = EditorGrid.Redraw
End Property

Public Property Let Redraw(ByVal vNewValue As Boolean)
On Error Resume Next
    EditorGrid.Redraw = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get RightToLeft() As Boolean
    RightToLeft = EditorGrid.RightToLeft
End Property

Public Property Let RightToLeft(ByVal vNewValue As Boolean)
On Error Resume Next
    EditorGrid.RightToLeft = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get RowHeightMin() As Long
    RowHeightMin = EditorGrid.RowHeightMin
End Property

Public Property Let RowHeightMin(ByVal vNewValue As Long)
On Error Resume Next
    EditorGrid.RowHeightMin = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get ScrollBars() As ScrollBarsSettings
    ScrollBars = EditorGrid.ScrollBars
End Property

Public Property Let ScrollBars(ByVal vNewValue As ScrollBarsSettings)
On Error Resume Next
    EditorGrid.ScrollBars = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get ScrollTrack() As Boolean
    ScrollTrack = EditorGrid.ScrollTrack
End Property

Public Property Let ScrollTrack(ByVal vNewValue As Boolean)
On Error Resume Next
    EditorGrid.ScrollTrack = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get SelectionMode() As SelectionModeSettings
    SelectionMode = EditorGrid.SelectionMode
End Property

Public Property Let SelectionMode(ByVal vNewValue As SelectionModeSettings)
On Error Resume Next
    EditorGrid.SelectionMode = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get TextStyle() As TextStyleSettings
    TextStyle = EditorGrid.TextStyle
End Property

Public Property Let TextStyle(ByVal vNewValue As TextStyleSettings)
On Error Resume Next
    EditorGrid.TextStyle = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get TextStyleFixed() As TextStyleSettings
    TextStyleFixed = EditorGrid.TextStyleFixed
End Property

Public Property Let TextStyleFixed(ByVal vNewValue As TextStyleSettings)
On Error Resume Next
    EditorGrid.TextStyleFixed = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get ToolTipText() As String
    ToolTipText = EditorGrid.ToolTipText
End Property

Public Property Let ToolTipText(ByVal vNewValue As String)
On Error Resume Next
    EditorGrid.ToolTipText = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = EditorGrid.WordWrap
End Property

Public Property Let WordWrap(ByVal vNewValue As Boolean)
On Error Resume Next
    EditorGrid.WordWrap = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get ColWidth(ByVal Col As Long) As Long
Attribute ColWidth.VB_MemberFlags = "400"
    ColWidth = EditorGrid.ColWidth(Col)
End Property

Public Property Let ColWidth(ByVal Col As Long, ByVal vNewValue As Long)
On Error Resume Next
    EditorGrid.ColWidth(Col) = vNewValue
    Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

Public Property Get MouseRow() As Long
Attribute MouseRow.VB_MemberFlags = "400"
    MouseRow = EditorGrid.MouseRow
End Property

Public Property Let MouseRow(ByVal vNewValue As Long)
    'EditorGrid.MouseRow = vNewValue
End Property
Public Property Get MouseCol() As Long
Attribute MouseCol.VB_MemberFlags = "400"
    MouseCol = EditorGrid.MouseCol
End Property

Public Property Let MouseCol(ByVal vNewValue As Long)
    'EditorGrid.MouseCol = vNewValue
End Property


Public Property Get RowHeight(gridCol As Long) As Long
Attribute RowHeight.VB_MemberFlags = "400"
    RowHeight = EditorGrid.RowHeight(gridCol)
End Property

Public Property Let RowHeight(gridCol As Long, ByVal vNewValue As Long)
On Error Resume Next
    EditorGrid.RowHeight(gridCol) = vNewValue
Exit Property
errHand:
MsgBox Err.Description, vbCritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Err.Clear
End Property

'Public Property Get dateVal() As Variant
'On Error resume next
'    dateVal = DatePicker
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Let dateVal(ByVal vNewValue As Variant)
'On Error resume next
'If Not IsNull(vNewValue) Then
'    DatePicker.Value = vNewValue
'Else
'    DatePicker.Value = Null
'End If
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Get MinDate() As Date
'On Error resume next
'    MinDate = DatePicker.MinDate
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Let MinDate(ByVal vNewValue As Date)
'On Error resume next
'    DatePicker.MinDate = vNewValue
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Get MaxDate() As Variant
'On Error resume next
'    MaxDate = DatePicker.MaxDate
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Let MaxDate(ByVal vNewValue As Variant)
'On Error resume next
'    DatePicker.MaxDate = vNewValue
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Get DateFormat() As Variant
'On Error resume next
'    DateFormat = DatePicker.Format
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Let DateFormat(ByVal vNewValue As Variant)
'On Error resume next
'    DatePicker.Format = vNewValue
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Get DateCustom() As Variant
'On Error resume next
'    DateCustom = DatePicker.CustomFormat
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Let DateCustom(ByVal vNewValue As Variant)
'On Error resume next
'    DatePicker.CustomFormat = vNewValue
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Get DateCheckbox() As Variant
'On Error resume next
'If DatePicker.CheckBox Then
'    DateCheckbox = 1
'Else: DateCheckbox = 0
'End If
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Let DateCheckbox(ByVal vNewValue As Variant)
'On Error resume next
'    DatePicker.CheckBox = vNewValue
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Get DateUpdown() As Variant
'On Error resume next
'If DatePicker.UpDown Then
'    DateUpdown = 1
'Else
'    DateUpdown = 0
'End If
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Let DateUpdown(ByVal vNewValue As Variant)
'On Error resume next
'    DatePicker.UpDown = vNewValue
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Get DateboxTime() As Variant
'On Error resume next
'    DateboxTime = TimePick
'Exit Property
'errHand:
'msgbox err.description,vbcritical 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'Err.Clear
'End Property
'
'Public Property Let DateboxTime(ByVal vNewValue As Variant)
'On Error Resume Next
'    TimePick.Value = vNewValue
'End Property

Public Property Get ColAlignment(Index As Long) As Variant
Attribute ColAlignment.VB_MemberFlags = "400"
On Error Resume Next
    ColAlignment = EditorGrid.ColAlignment(Index)
End Property

Public Property Let ColAlignment(Index As Long, ByVal vNewValue As Variant)
On Error Resume Next
    EditorGrid.ColAlignment(Index) = vNewValue
End Property

Public Property Get CellAlignment() As Variant
Attribute CellAlignment.VB_MemberFlags = "400"
On Error Resume Next
    CellAlignment = EditorGrid.CellAlignment
End Property

Public Property Let CellAlignment(ByVal vNewValue As Variant)
On Error Resume Next
    EditorGrid.CellAlignment = vNewValue
End Property

Public Property Get EditorTime() As Variant
Attribute EditorTime.VB_MemberFlags = "400"
On Error Resume Next
    EditorTime = TimePicker.Value
End Property

Public Property Let EditorTime(ByVal vNewValue As Variant)
On Error Resume Next
    TimePicker = TimeValue(Format(vNewValue, "hh:mm:ss"))
End Property

Public Property Get EditorDate() As Variant
Attribute EditorDate.VB_MemberFlags = "400"
On Error Resume Next
EditorDate = DatePicker
End Property

Public Property Let EditorDate(ByVal vNewValue As Variant)
On Error Resume Next
    DatePicker = DateValue(Format(vNewValue, "dd/mm/yyyy"))
End Property

Public Property Get CreateNewRows() As Boolean
On Error Resume Next
    CreateNewRows = CreateNewRows1
End Property

Public Property Let CreateNewRows(ByVal vNewValue As Boolean)
    CreateNewRows1 = vNewValue
End Property

Public Property Get CellFontBold() As Boolean
Attribute CellFontBold.VB_MemberFlags = "400"
    CellFontBold = EditorGrid.CellFontBold
End Property

Public Property Let CellFontBold(ByVal vNewValue As Boolean)
    EditorGrid.CellFontBold = vNewValue
End Property


Public Property Get CellBackColor() As OLE_COLOR
Attribute CellBackColor.VB_MemberFlags = "400"
    CellBackColor = EditorGrid.CellBackColor
End Property

Public Property Let CellBackColor(ByVal vNewValue As OLE_COLOR)
On Error Resume Next
    EditorGrid.CellBackColor = vNewValue
End Property
Public Property Get CellforeColor() As OLE_COLOR
Attribute CellforeColor.VB_MemberFlags = "400"
    CellforeColor = EditorGrid.CellforeColor
End Property

Public Property Let CellforeColor(ByVal vNewValue As OLE_COLOR)
On Error Resume Next
    EditorGrid.CellforeColor = vNewValue
End Property

Public Sub ComboAdditem(Item As Variant)
On Error Resume Next
    CmbEditor.AddItem Item
End Sub

Public Property Get ComboDropButtonStyle() As fmDropButtonStyle
    ComboDropButtonStyle = CmbEditor.DropButtonStyle
End Property

Public Property Let ComboDropButtonStyle(ByVal vNewValue As fmDropButtonStyle)
On Error Resume Next
    CmbEditor.DropButtonStyle = vNewValue
End Property

Public Property Get ComboAutoWordSelect() As Boolean
    ComboAutoWordSelect = CmbEditor.AutoWordSelect
End Property

Public Property Let ComboAutoWordSelect(ByVal vNewValue As Boolean)
    CmbEditor.AutoWordSelect = vNewValue
    PropertyChanged "ComboAutoWordSelect"
End Property

Public Property Get ComboMachEntry() As fmMatchEntry
Attribute ComboMachEntry.VB_MemberFlags = "400"
    ComboMachEntry = CmbEditor.MatchEntry
End Property

Public Property Let ComboMachEntry(ByVal vNewValue As fmMatchEntry)
    CmbEditor.MatchEntry = vNewValue
    PropertyChanged "ComboMachEntry"
End Property

Public Property Get ComboStyle() As fmStyle
    ComboStyle = CmbEditor.Style
End Property

Public Property Let ComboStyle(ByVal vNewValue As fmStyle)
On Error Resume Next
    CmbEditor.Style = vNewValue
    PropertyChanged "ComboStyle"
End Property

Public Property Get SetMaxLength(ByVal Column As Long) As Long
Attribute SetMaxLength.VB_MemberFlags = "400"
End Property

Public Property Let SetMaxLength(ByVal Column As Long, ByVal vNewValue As Long)
Dim i, j, Col1 As Long
    With MaxL
        For i = 0 To .ListCount - 1
            For j = 1 To Len(.List(i))
                If Mid(.List(i), j, 1) = ";" Then
                    Col1 = Left(.List(i), j - 1)
                    If Col1 = Column Then
                        .RemoveItem i
                    End If
                End If
            Next
        Next
    End With
    MaxL.AddItem Column & ";" & vNewValue
End Property

Public Property Get CalenderBackColor() As OLE_COLOR
    CalenderBackColor = DatePicker.CalendarBackColor
End Property

Public Property Let CalenderBackColor(ByVal vNewValue As OLE_COLOR)
    DatePicker.CalendarBackColor = vNewValue
End Property

Public Property Get CalenderForeColor() As OLE_COLOR
    CalenderForeColor = DatePicker.CalendarForeColor
End Property

Public Property Let CalenderForeColor(ByVal vNewValue As OLE_COLOR)
    DatePicker.CalendarForeColor = vNewValue
End Property

Public Property Get CalenderTitleBackColor() As OLE_COLOR
    CalenderTitleBackColor = DatePicker.CalendarTitleBackColor
End Property

Public Property Let CalenderTitleBackColor(ByVal vNewValue As OLE_COLOR)
DatePicker.CalendarTitleBackColor = vNewValue
End Property

Public Property Get CalenderTitleForeColor() As OLE_COLOR
    CalenderTitleForeColor = DatePicker.CalendarTitleForeColor
End Property

Public Property Let CalenderTitleForeColor(ByVal vNewValue As OLE_COLOR)
    DatePicker.CalendarTitleForeColor = vNewValue
End Property

Public Property Get CalenderTrailingForeColor() As OLE_COLOR
    CalenderTrailingForeColor = DatePicker.CalendarTrailingForeColor
End Property

Public Property Let CalenderTrailingForeColor(ByVal vNewValue As OLE_COLOR)
    DatePicker.CalendarTrailingForeColor = vNewValue
End Property

Public Property Get CalenderFormat() As FormatConstants
    CalenderFormat = DatePicker.Format
End Property

Public Property Let CalenderFormat(ByVal vNewValue As FormatConstants)
    DatePicker.Format = vNewValue
End Property

Public Property Get CalenderCustomFormat() As String
    CalenderCustomFormat = DatePicker.CustomFormat
End Property

Public Property Let CalenderCustomFormat(ByVal vNewValue As String)
    DatePicker.CustomFormat = vNewValue
End Property

Public Property Get CalenderUpdown() As Boolean
    CalenderUpdown = DatePicker.UpDown
End Property

Public Property Let CalenderUpdown(ByVal vNewValue As Boolean)
    DatePicker.UpDown = vNewValue
End Property
Public Property Get NumbersOnly(ByVal Col As Long, Optional ByVal Value As Integer = 2, Optional ByVal Limit As Double = 99999.99) As Boolean
Attribute NumbersOnly.VB_MemberFlags = "400"
    
End Property

Public Property Let NumbersOnly(ByVal Col As Long, Optional ByVal Value As Integer = 2, Optional ByVal Limit As Double = 99999.99, ByVal vNewValue As Boolean)
'To enter only numbers in a particular cell
Dim i, j, Col1 As Long
    With OnlyNo
        For i = 0 To .ListCount - 1
            For j = 1 To Len(.List(i))
                If Mid(.List(i), j, 1) = ";" Then
                    Col1 = Val(Left(.List(i), j - 1))
                    If Col1 = Col Then 'If the value is already there then remove.
                        OnlyNo.RemoveItem i
                    End If
                End If
            Next
        Next
    End With
    OnlyNo.AddItem Col & ";" & vNewValue & ";" & Value & ";" & Limit & ";"
End Property


Public Property Get CellPicture() As IPictureDisp
Attribute CellPicture.VB_MemberFlags = "400"
    
End Property

Public Property Set CellPicture(ByVal vNewValue As IPictureDisp)
On Error Resume Next
    Set EditorGrid.CellPicture = vNewValue
End Property

Public Property Get CellPictureAlignment() As Integer
Attribute CellPictureAlignment.VB_MemberFlags = "400"
    
End Property

Public Property Let CellPictureAlignment(ByVal vNewValue As Integer)
On Error Resume Next
    EditorGrid.CellPictureAlignment = vNewValue
End Property

'Public Property Get EditorTimeFormat() As String
'    On Error Resume Next
'        EditorTimeFormat = TimePicker.CustomFormat
'End Property
'
'Public Property Let EditorTimeFormat(ByVal vNewValue As String)
'    On Error Resume Next
'    TimePicker.CustomFormat = vNewValue
'End Property

Public Sub ComboDropDown()
On Error Resume Next
    CmbEditor.DropDown
End Sub

Public Sub Refresh()
On Error Resume Next
    EditorGrid.Refresh
End Sub

'Public Sub Drag()
'    EditorGrid.Drag
'End Sub
'
'Public Property Get DragIcon() As StdPicture
'
'End Property
'
'Public Property Set DragIcon(ByVal vNewValue As StdPicture)
'On Error Resume Next
'    Set EditorGrid.DragIcon = vNewValue
'End Property

Public Property Get CellTop() As Variant
Attribute CellTop.VB_MemberFlags = "400"
    CellTop = EditorGrid.CellTop
End Property

Public Property Let CellTop(ByVal vNewValue As Variant)

End Property

Public Property Get CellHeight() As Variant
Attribute CellHeight.VB_MemberFlags = "400"
    CellHeight = EditorGrid.CellHeight
End Property

Public Property Let CellHeight(ByVal vNewValue As Variant)

End Property

Public Property Get CellWidth() As Variant
Attribute CellWidth.VB_MemberFlags = "400"
    CellWidth = EditorGrid.CellWidth
End Property

Public Property Let CellWidth(ByVal vNewValue As Variant)

End Property

Public Property Get CellLeft() As Variant
Attribute CellLeft.VB_MemberFlags = "400"
    CellLeft = EditorGrid.CellLeft
End Property

Public Property Let CellLeft(ByVal vNewValue As Variant)

End Property

Public Property Get CellTextStyle() As TextStyleSettings
Attribute CellTextStyle.VB_MemberFlags = "400"
    CellTextStyle = EditorGrid.CellTextStyle
End Property

Public Property Let CellTextStyle(ByVal vNewValue As TextStyleSettings)
    EditorGrid.CellTextStyle = vNewValue
End Property

Public Property Get ColData(Index As Long) As Long
Attribute ColData.VB_MemberFlags = "400"
    ColData = EditorGrid.ColData(Index)
End Property

Public Property Let ColData(Index As Long, ByVal vNewValue As Long)
    EditorGrid.ColData(Index) = vNewValue
End Property

Public Property Get ColIsVisible(Index As Long) As Boolean
Attribute ColIsVisible.VB_MemberFlags = "400"
    ColIsVisible = EditorGrid.ColIsVisible(Index)
End Property

Public Property Let ColIsVisible(Index As Long, ByVal vNewValue As Boolean)

End Property

Public Property Get ColPos(Index As Long) As Long
Attribute ColPos.VB_MemberFlags = "400"
    ColPos = EditorGrid.ColPos(Index)
End Property

Public Property Let ColPos(Index As Long, ByVal vNewValue As Long)
    
End Property

Public Property Get ColSel() As Long
Attribute ColSel.VB_MemberFlags = "400"
    ColSel = EditorGrid.ColSel
End Property

Public Property Let ColSel(ByVal vNewValue As Long)
    EditorGrid.ColSel = vNewValue
End Property

Public Sub ClearCombo()
On Error Resume Next
    CmbEditor.Clear
End Sub

Public Sub ComboRemoveItem(Index As Long)
On Error Resume Next
    CmbEditor.RemoveItem Index
End Sub


Public Property Get ComboMatchFound() As Boolean
Attribute ComboMatchFound.VB_MemberFlags = "400"
On Error Resume Next
    ComboMatchFound = CmbEditor.MatchFound
End Property

Public Property Let ComboMatchFound(ByVal vNewValue As Boolean)
End Property

Public Property Get ComboMatchRequired() As Boolean
    ComboMatchRequired = CmbEditor.MatchRequired
End Property

Public Property Let ComboMatchRequired(ByVal vNewValue As Boolean)
On Error Resume Next
    CmbEditor.MatchRequired = vNewValue
End Property

Public Property Get ComboShowDropButtonWhen() As fmShowDropButtonWhen
On Error Resume Next
    ComboShowDropButtonWhen = CmbEditor.ShowDropButtonWhen
End Property

Public Property Let ComboShowDropButtonWhen(ByVal vNewValue As fmShowDropButtonWhen)
On Error Resume Next
    CmbEditor.ShowDropButtonWhen = vNewValue
End Property

Public Sub GenerateGridNumber()
On Error Resume Next
Dim i As Long
With EditorGrid
    For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = i
    Next
End With
End Sub

Public Property Get MergeRow(Index As Long) As Boolean
Attribute MergeRow.VB_MemberFlags = "400"
On Error Resume Next
MergeRow = EditorGrid.MergeRow(Index)
End Property

Public Property Let MergeRow(Index As Long, ByVal vNewValue As Boolean)
On Error Resume Next
EditorGrid.MergeRow(Index) = vNewValue
End Property

Public Property Get MergeCol(Index As Long) As Boolean
Attribute MergeCol.VB_MemberFlags = "400"
On Error Resume Next
MergeCol = EditorGrid.MergeCol(Index)
End Property

Public Property Let MergeCol(Index As Long, ByVal vNewValue As Boolean)
On Error Resume Next
EditorGrid.MergeCol(Index) = vNewValue
End Property

Public Property Get Enabled() As Boolean
On Error Resume Next
Enabled = EditorGrid.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
On Error Resume Next
EditorGrid.Enabled = vNewValue
End Property

Public Property Get AllowUserResizing() As AllowUserResizeSettings
On Error Resume Next
AllowUserResizing = EditorGrid.AllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal vNewValue As AllowUserResizeSettings)
EditorGrid.AllowUserResizing = vNewValue
On Error Resume Next
End Property

Public Property Get ComboIndex() As Long
Attribute ComboIndex.VB_MemberFlags = "400"
ComboIndex = CmbEditor.ListIndex
End Property

Public Property Let ComboIndex(ByVal vNewValue As Long)
CmbEditor.ListIndex = vNewValue
End Property

Public Property Get ComboListCount() As Long
Attribute ComboListCount.VB_MemberFlags = "400"
ComboListCount = CmbEditor.ListCount
End Property

'--------------------------Added on 4/2/02
Public Property Get Sort() As Integer
Attribute Sort.VB_MemberFlags = "400"
    'Sort = EditorGrid.Sort
End Property

Public Property Let Sort(ByVal vNewValue As Integer)
    EditorGrid.Sort = vNewValue
End Property

Public Property Get Appearance() As AppearanceSettings
    Appearance = EditorGrid.Appearance
End Property

Public Property Let Appearance(ByVal vNewValue As AppearanceSettings)
    EditorGrid.Appearance = vNewValue
End Property

Public Property Get CellFontItalic() As Boolean
Attribute CellFontItalic.VB_MemberFlags = "400"
    CellFontItalic = EditorGrid.CellFontItalic
End Property

Public Property Let CellFontItalic(ByVal vNewValue As Boolean)
    EditorGrid.CellFontItalic = vNewValue
End Property

Public Property Get CellFontName() As String
Attribute CellFontName.VB_MemberFlags = "400"
    CellFontName = EditorGrid.CellFontName
End Property

Public Property Let CellFontName(ByVal vNewValue As String)
    EditorGrid.CellFontName = vNewValue
End Property

Public Property Get CellFontSize() As Single
Attribute CellFontSize.VB_MemberFlags = "400"
    CellFontSize = EditorGrid.CellFontSize
End Property

Public Property Let CellFontSize(ByVal vNewValue As Single)
    EditorGrid.CellFontSize = vNewValue
End Property

Public Property Get CellFontStrikeThrough() As Boolean
Attribute CellFontStrikeThrough.VB_MemberFlags = "400"
    CellFontStrikeThrough = EditorGrid.CellFontStrikeThrough
End Property

Public Property Let CellFontStrikeThrough(ByVal vNewValue As Boolean)
    EditorGrid.CellFontStrikeThrough = vNewValue
End Property

Public Property Get CellFontUnderLine() As Boolean
Attribute CellFontUnderLine.VB_MemberFlags = "400"
    CellFontUnderLine = EditorGrid.CellFontUnderLine
End Property

Public Property Let CellFontUnderLine(ByVal vNewValue As Boolean)
    EditorGrid.CellFontUnderLine = vNewValue
End Property

Public Property Get CellfontWidth() As Single
Attribute CellfontWidth.VB_MemberFlags = "400"
    CellfontWidth = EditorGrid.CellfontWidth
End Property

Public Property Let CellfontWidth(ByVal vNewValue As Single)
    EditorGrid.CellfontWidth = vNewValue
End Property

Public Property Get Clip() As String
Attribute Clip.VB_MemberFlags = "400"
    Clip = EditorGrid.Clip
End Property

Public Property Let Clip(ByVal vNewValue As String)
    EditorGrid.Clip = vNewValue
End Property

Public Property Get FixedAlignment(Index As Long) As AlignmentConstants
    FixedAlignment = EditorGrid.FixedAlignment(Index)
End Property

Public Property Let FixedAlignment(Index As Long, ByVal vNewValue As AlignmentConstants)
    EditorGrid.FixedAlignment(Index) = vNewValue
End Property

Public Property Get FontWidth() As Single
    FontWidth = EditorGrid.FontWidth
End Property

Public Property Let FontWidth(ByVal vNewValue As Single)
    EditorGrid.FontWidth = vNewValue
End Property

Public Property Get LeftCol() As Long
Attribute LeftCol.VB_MemberFlags = "400"
    LeftCol = EditorGrid.LeftCol
End Property

Public Property Let LeftCol(ByVal vNewValue As Long)
    EditorGrid.LeftCol = vNewValue
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_MemberFlags = "400"
    Set Picture = EditorGrid.Picture
End Property

Public Property Set Picture(ByVal vNewValue As IPictureDisp)
'    Set EditorGrid.Picture = vNewValue
End Property

Public Property Get RowData(Index As Long) As Long
Attribute RowData.VB_MemberFlags = "400"
    RowData = EditorGrid.RowData(Index)
End Property

Public Property Let RowData(Index As Long, ByVal vNewValue As Long)
    EditorGrid.RowData(Index) = vNewValue
End Property

Public Property Get RowIsVisible(Index As Long) As Boolean
Attribute RowIsVisible.VB_MemberFlags = "400"
    RowIsVisible = EditorGrid.RowIsVisible(Index)
End Property

Public Property Get RowPos(Index As Long) As Long
Attribute RowPos.VB_MemberFlags = "400"
    RowPos = EditorGrid.RowPos(Index)
End Property

Public Property Get RowPosition(Index As Long) As Long
Attribute RowPosition.VB_MemberFlags = "400"
'    RowPosition = EditorGrid.RowPosition(Index)
End Property

Public Property Get RowSel() As Variant
Attribute RowSel.VB_MemberFlags = "400"
    RowSel = EditorGrid.RowSel
End Property

Public Property Let RowSel(ByVal vNewValue As Variant)
    EditorGrid.RowSel = vNewValue
End Property

Public Property Get TextArray(Index As Long) As Variant
Attribute TextArray.VB_MemberFlags = "400"
    TextArray = EditorGrid.TextArray(Index)
End Property

Public Property Let TextArray(Index As Long, ByVal vNewValue As Variant)
    EditorGrid.TextArray(Index) = vNewValue
End Property

Public Property Get TopRow() As Long
Attribute TopRow.VB_MemberFlags = "400"
    TopRow = EditorGrid.TopRow
End Property
Public Property Let TopRow(vNewValue As Long)
    EditorGrid.TopRow = vNewValue
End Property


'--------------------------End of Added on 4/2/02

'**********************************************************************
'Functions used for Print Grid to a printer

Public Sub PrintGrid()
On Error Resume Next
Dim CreateNRow As Boolean
CreateNRow = CreateNewRows1
CreateNewRows1 = False
showprintForm
NoofPages1 = 1
Screen.MousePointer = vbHourglass
    Call InitializePictureBox
    Set cTP = New clsTablePrint
    Call CmdRefresh_Click
    Call cmdPrint_Click
Screen.MousePointer = vbDefault
Picture1.Visible = False
CreateNewRows1 = CreateNRow
End Sub
Private Sub showprintForm()
On Error Resume Next
With Picture1
    .Left = UserControl.ScaleWidth / 2 - .Width / 2
    .Top = UserControl.ScaleHeight / 2 - .Height / 2
    .Visible = True
    UserControl.Refresh
End With
End Sub
Public Sub PrintPreview()
Attribute PrintPreview.VB_MemberFlags = "40"

End Sub
Private Sub cmdPrint_Click()
Dim CreateNRow As Boolean
    CreateNRow = CreateNewRows1
    CreateNewRows1 = False

    'If MsgBox("The application will now print the grid on the default printer (Show a print dialog here later !).", vbInformation + vbOKCancel, "Print") = vbCancel Then Exit Sub
    
    'Simply initialize the printer:
    Printer.Orientation = 2
    Printer.Print
    
    'Read the FlexGrid:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    ImportFlexGrid cTP, EditorGrid, IIf((chkColWidth.Value = vbChecked), Printer.ScaleWidth - 2 * 567, -1)
    
    'Set margins (not needed, but looks better !):
    cTP.MarginBottom = 567 '567 equals to 1 cm
    cTP.MarginLeft = 567
    cTP.MarginTop = 567
    
    'Class begins drawing at CurrentY !
    Printer.CurrentY = cTP.MarginTop
    
    'Finally draw the Grid !
    
    cTP.DrawTable Printer
    'Done with drawing !
    
    'Say VB it should finally send it:
    Printer.EndDoc
    CreateNewRows1 = CreateNRow
End Sub

Private Sub InitializePictureBox()
    Dim sngVSCWidth As Single, sngHSCHeight As Single
    'Set the size to the DIN A4 width:
    picTarget.Width = A4Width
    picTarget.Height = A4Height
    'Resize the scrollbars:
    sngVSCWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    sngHSCHeight = GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY
    hscScroll.Move 0, picScroll.ScaleHeight - sngHSCHeight, picScroll.ScaleWidth - sngVSCWidth, sngHSCHeight
    vscScroll.Move picScroll.ScaleWidth - sngVSCWidth, 0, sngVSCWidth, picScroll.ScaleHeight
    
    SetScrollBars
End Sub
Private Sub InitializePictureBoxToPreview()
    Dim sngVSCWidth As Single, sngHSCHeight As Single
    'Set the size to the DIN A4 width:
    With Frm_Grid_Preview
        .picTarget.Width = A4Width
        .picTarget.Height = A4Height
        .Width = A4Width
        .Height = A4Height
        'Resize the scrollbars:
        sngVSCWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
        sngHSCHeight = GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY
        .hscScroll.Move 0, .picScroll.ScaleHeight - sngHSCHeight, .picScroll.ScaleWidth - sngVSCWidth, sngHSCHeight
        .vscScroll.Move .picScroll.ScaleWidth - sngVSCWidth, 0, sngVSCWidth, .picScroll.ScaleHeight
        
        SetScrollBarsToPreview
    End With
End Sub
Private Sub CmdRefresh_Click()
Dim CreateNRow As Boolean
    CreateNRow = CreateNewRows1
    CreateNewRows1 = False

    'Read the FlexGrid:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    ImportFlexGrid cTP, EditorGrid, IIf((chkColWidth.Value = vbChecked), picTarget.ScaleWidth - 2 * 567, -1)
    
    'Set margins (not needed, but looks better !):
    cTP.MarginBottom = 567 '567 equals to 1 cm
    cTP.MarginLeft = 567
    cTP.MarginTop = 567
    
    'Clear the box:
    picTarget.Cls
    
    'Class begins drawing at CurrentY !
    picTarget.CurrentY = cTP.MarginTop
    
    'Finally draw the Grid !
    cTP.DrawTable picTarget
    'Done with drawing !
    CreateNewRows1 = CreateNRow
End Sub
Private Sub cmdRefresh_ClicktoPreview()
Dim CreateNRow As Boolean
CreateNRow = CreateNewRows1
CreateNewRows1 = False

    'Read the FlexGrid:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    With Frm_Grid_Preview
    ImportFlexGrid cTP, EditorGrid, IIf((chkColWidth.Value = vbChecked), .picTarget.ScaleWidth - 2 * 567, -1)
    
    'Set margins (not needed, but looks better !):
    cTP.MarginBottom = 567 '567 equals to 1 cm
    cTP.MarginLeft = 567
    cTP.MarginTop = 567

    'Clear the box:
    .picTarget.Cls
    
    'Class begins drawing at CurrentY !
    .picTarget.CurrentY = cTP.MarginTop
    
    'Finally draw the Grid !
    cTP.DrawTable .picTarget
    'Done with drawing !
End With
CreateNewRows1 = CreateNRow
End Sub

Private Sub SetScrollBarsToPreview()
'With Frm_Grid_Preview
'    .hscScroll.Max = (.picTarget.Width - .picScroll.ScaleWidth + .vscScroll.Width) / 120 + 1
'    .vscScroll.Max = (.picTarget.Height - .picScroll.ScaleHeight + .hscScroll.Height) / 120 + 1
'    '.hscScroll.Top = .picTarget.ScaleHeight - .hscScroll.Height - 300
'
'    .hscScroll.Width = .picTarget.ScaleWidth
'    .hscScroll.Left = 0 '.picTarget.ScaleWidth - .hscScroll.Height
'    '.hscScroll.Height = .picTarget.ScaleHeight
'    .hscScroll.Top = .picTarget.ScaleHeight - .hscScroll.Height - 300
'
'    .vscScroll.Top = .picTarget.ScaleHeight - .vscScroll.Height - 300
'    .vscScroll.Width = .picTarget.ScaleWidth
'    .vscScroll.Left = 0
'End With
End Sub
Private Sub SetScrollBars()
    hscScroll.Max = (picTarget.Width - picScroll.ScaleWidth + vscScroll.Width) / 120 + 1
    vscScroll.Max = (picTarget.Height - picScroll.ScaleHeight + hscScroll.Height) / 120 + 1
End Sub
'*******************************End of functions used for PrintGrid***************************************
Private Sub vscScroll_Change()
picTarget.Top = -CSng(vscScroll.Value) * 120
End Sub

Public Property Get NoofPages() As Long
    NoofPages = NoofPages1
End Property
Public Sub PrintView(Optional Fullscreen As Boolean = False)
On Error Resume Next
Dim CreateNRow As Boolean
CreateNRow = CreateNewRows1
CreateNewRows1 = False
If Fullscreen Then Call CmdFull_Click
    Screen.MousePointer = vbHourglass
    EditorGrid.Visible = False
    NoofPages1 = 1
    Call InitializePictureBox
    Set cTP = New clsTablePrint
    Call CmdRefresh_Click
Screen.MousePointer = vbDefault
CreateNewRows1 = CreateNRow
picScroll.Visible = True
Cmd_Print.Visible = True
Cmd_Close.Visible = True
chkColWidth.Visible = True
CmdFull.Visible = True
Call UserControl_Resize
End Sub

Public Property Get EditorBackColor() As OLE_COLOR
    EditorBackColor = TxtEditor.BackColor
End Property

Public Property Let EditorBackColor(ByVal vNewValue As OLE_COLOR)
    TxtEditor.BackColor = vNewValue
    CmbEditor.BackColor = vNewValue
    EditorColor1 = vNewValue
    'DatePicker.BackColor = vNewValue
'    TimePicker.b
End Property
Public Property Get EditorForeColor() As OLE_COLOR
    EditorForeColor = TxtEditor.ForeColor
End Property

Public Property Let EditorForeColor(ByVal vNewValue As OLE_COLOR)
    TxtEditor.ForeColor = vNewValue
    CmbEditor.ForeColor = vNewValue
    EditorForeColor1 = vNewValue
'    DatePicker.BackColor = vNewValue
'    TimePicker.b
End Property


Public Property Get EnableCellEdit(Row As Long, Col As Long) As Boolean
On Error Resume Next
If Not CheckEnableEdit Then EnableCellEdit = True
CheckEnableEdit = False
Dim j As Long
Dim CellArray() As String
Dim flag As Boolean
For j = 0 To UBound(CellEditArray)
    CellArray = Split(CellEditArray(j), ";")
    If Not (UBound(CellArray) = 0 Or UBound(CellArray) = -1) Then
        If Val(CellArray(0)) = Row And Val(CellArray(1)) = Col Then
            If Val(CellArray(2)) = 1 Then
                EnableCellEdit = True
            Else
                EnableCellEdit = False
            End If
            Exit For
        End If
    End If
Next
End Property

Public Property Let EnableCellEdit(Row As Long, Col As Long, ByVal vNewValue As Boolean)
On Error Resume Next
Dim i As Integer
Dim j As Long
Dim CellArray() As String
Dim flag As Boolean
If vNewValue Then i = 1

For j = 0 To UBound(CellEditArray)
    CellArray = Split(CellEditArray(j), ";")
    If Not (UBound(CellArray) = 0 Or UBound(CellArray)) = -1 Then
        If Val(CellArray(0)) = Row And Val(CellArray(1)) = Col Then
            CellEditArray(i) = Row & ";" & Col & ";" & i
            flag = True
            Exit For
        End If
    End If
Next
If Not flag Then
    ReDim Preserve CellEditArray(0 To UBound(CellEditArray) + 1)
    CellEditArray(UBound(CellEditArray)) = Row & ";" & Col & ";" & i
End If
End Property

Public Property Get EnableRowEdit(Row As Long) As Boolean
On Error Resume Next
EnableRowEdit = True
Dim j As Long
Dim CellArray() As String
Dim flag As Boolean
For j = 0 To UBound(RowEditArray)
    CellArray = Split(RowEditArray(j), ";")
    If Not (UBound(CellArray) = 0 Or UBound(CellArray) = -1) Then
        If Val(CellArray(0)) = Row Then
            If Val(CellArray(1)) = 1 Then
                EnableRowEdit = True
            Else
                EnableRowEdit = False
            End If
            Exit For
        End If
    End If
Next
End Property

Public Property Let EnableRowEdit(Row As Long, ByVal vNewValue As Boolean)
On Error Resume Next
Dim i As Integer
Dim j As Long
Dim CellArray() As String
Dim flag As Boolean
If vNewValue Then i = 1
For j = 0 To UBound(RowEditArray)
    CellArray = Split(RowEditArray(j), ";")
    If Not (UBound(CellArray) = 0 Or UBound(CellArray)) = -1 Then
        If Val(CellArray(0)) = Row Then
            RowEditArray(i) = Row & ";" & i
            flag = True
            Exit For
        End If
    End If
Next
If Not flag Then
    ReDim Preserve RowEditArray(0 To UBound(RowEditArray) + 1)
    RowEditArray(UBound(RowEditArray)) = Row & ";" & i
End If
End Property

Public Property Get PrintOrientation() As PrinterOrientation
    PrintOrientation = PrintOrient
End Property

Public Property Let PrintOrientation(ByVal vNewValue As PrinterOrientation)
    PrintOrient = vNewValue
End Property

Public Sub SaveGrid(Filename As String)
On Error GoTo errH
Dim Str As String
With EditorGrid
    If Not (.Cols > 0 And .Rows > 0) Then Exit Sub
    .Col = 0: .Row = 0
    .ColSel = .Cols - 1
    .RowSel = .Rows - 1
    Str = .Clip
    Open Filename For Output As #1
    Print #1, Str
    Close #1
End With
Exit Sub
errH:
Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
