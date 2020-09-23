VERSION 5.00
Object = "*\AGridEditorControl.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_Gridsample 
   Caption         =   "Grid Editor Sample"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4170
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ActifOcx.GridEditor GridEditor1 
      Height          =   3570
      Left            =   90
      TabIndex        =   3
      Top             =   510
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   6297
      ROWS            =   20
      COLS            =   10
      BeginProperty FONT1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BACKCOLOR       =   -2147483643
      BACKCOLORBKG    =   8421504
      BACKCOLORFIXED  =   -2147483633
      BACKCOLORSEL    =   -2147483635
      FILLSTYLE       =   0
      FIXEDCOLS       =   1
      FIXEDROWS       =   1
      FOCUSRECT       =   1
      FORECOLOR       =   -2147483640
      FORECOLORFIXED  =   -2147483630
      FORECOLORSEL    =   -2147483634
      FORMATSTRING    =   ""
      GRIDCOLOR       =   12632256
      GRIDCOLORFIXED  =   0
      GRIDLINES       =   1
      GRIDLINESFIXED  =   2
      GRIDLINEWIDTH   =   1
      MERGECELLS      =   0
      MOUSEICON       =   "Frm_GridEditor_Sample.frx":0000
      MOUSEPOINTER    =   0
      PICTURETYPE     =   0
      REDRAW          =   -1  'True
      RIGHTTOLEFT     =   0   'False
      ROWHEIGHTMIN    =   285
      SCROLLBARS      =   3
      SCROLLTRACK     =   0   'False
      SELECTIONMODE   =   0
      TEXTSTYLE       =   0
      TEXTSTYLEFIXED  =   0
      Object.TOOLTIPTEXT     =   ""
      WORDWRAP        =   -1  'True
      CREATENEWROWS   =   -1  'True
      ComboAutoWordSelect=   -1  'True
      ComboMachEntry  =   1
      ComboStyle      =   0
      CALENDERBACKCOLOR=   -2147483643
      CALENDERFORECOLOR=   -2147483630
      CALENDERTITLEBACKCOLOR=   -2147483633
      CALENDERTITLEFORECLOR=   -2147483630
      CALENDERTRAILING=   -2147483631
      CALENDERFORMAT  =   1
      CALNDERCUSTOMFORMAT=   ""
      CALENDERUPDOWN  =   0   'False
      COMBOMATCHREQUIRED=   0   'False
      SHOWDROPBUTTONWHEN=   2
      ALLOWUSERRESIZING=   0
      ENABLED         =   -1  'True
      PrintOrient     =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CheckBox Demo"
      Height          =   390
      Left            =   10365
      TabIndex        =   2
      Top             =   0
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1515
      Top             =   15
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   45
      TabIndex        =   1
      Top             =   360
      Width           =   11685
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Editor Sample Application."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3660
      TabIndex        =   0
      Top             =   30
      Width           =   4650
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Editor Sample Application."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   270
      Left            =   3690
      TabIndex        =   4
      Top             =   45
      Width           =   4650
   End
   Begin VB.Menu Opt 
      Caption         =   "Options"
      Begin VB.Menu Opt1 
         Caption         =   "Show Editor"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu Opt1 
         Caption         =   "Allow Create New Rows"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu Opt1 
         Caption         =   "Set Calender Custom Format"
         Index           =   2
      End
      Begin VB.Menu Opt1 
         Caption         =   "Set Editor Date"
         Index           =   3
      End
      Begin VB.Menu Opt1 
         Caption         =   "Set Editor Time"
         Index           =   4
      End
      Begin VB.Menu Opt1 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu Opt1 
         Caption         =   "Set Editor Color"
         Index           =   6
      End
      Begin VB.Menu Opt1 
         Caption         =   "Set Editor ForeColor"
         Index           =   7
      End
      Begin VB.Menu Opt1 
         Caption         =   "Set Editor Font"
         Index           =   8
      End
      Begin VB.Menu Opt1 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu Opt1 
         Caption         =   "Print preview"
         Index           =   10
      End
      Begin VB.Menu Opt1 
         Caption         =   "Print"
         Index           =   11
      End
   End
   Begin VB.Menu Abt 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Frm_Gridsample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loading As Boolean

Private Sub SetGrid1()
Dim i As Long
    With GridEditor1
        .ShowEditor = True 'Shows the Editor
        .CreateNewRows = True 'Creates new row automatically when the user types in the last row
        .Row = 0 'Set Row to 0 .This is to set the column ca
        For i = 1 To .Cols - 1
            .Col = i
            .CellAlignment = 4 'Set captions to Center-Center
        Next
        .TextMatrix(0, 1) = "Text Box": .TextMatrix(0, 2) = "DateBox": .TextMatrix(0, 3) = "Combobox":
        .TextMatrix(0, 4) = "Disabled": .TextMatrix(0, 5) = "Time Box": .TextMatrix(0, 6) = "Number Box":
        .TextMatrix(0, 7) = "Number Box": .TextMatrix(0, 8) = "Text Box": .TextMatrix(0, 9) = "Text Box":
        .ColWidth(1) = 1700: .ColWidth(0) = 300: .ColWidth(2) = 1800: .ColWidth(5) = 1500
        For i = 1 To .Rows - 1
            .RowHeight(i) = 350
        Next
        .RowHeight(0) = 500
        For i = 1 To .Cols - 1
            .Col = i
            .CellAlignment = 4 'Set Alignment to Center - Center
            .CellFontBold = True
        Next
'---------------Sets Each grid column Objects
'if you dont specify any object for a particular cell then default object is a text box
        .SetColObject(2) = DateBox 'This property sets the col 2 to datebox
        .SetColObject(3) = ComboBox 'This property sets the col 3 to textbox
        .SetColObject(5) = TimeBox 'This property sets the col 5 to timebox
        '.SetColObject(6) = CheckBox 'This property sets the col 6 to Checkbox
        
        .EnableEditing(4) = False 'This property disables the editing propery of col 4
'---------------Adding Data to Combobox
        .ComboAdditem "Computer" 'Adds to the combo box
        .ComboAdditem "Mouse"
        .ComboAdditem "Keyboard"
        .ComboAdditem "Books"
        .ComboAdditem "Monitor"
        .ComboAdditem "Sound Card"
        .ComboAdditem "Hard disk"
        
'----------------This property numbersonly is used to restrict the user to enter only numbers in a particular grid
'----------------Format NumbersOnly(Cell No.,[No.of Decimal Places],[Maximum Value])
        
        .NumbersOnly(6, 2, 1000.99) = True 'This will restrict the user to enter only upto 1000.99 for the column 6
        .NumbersOnly(7, 0, 10) = True 'This will restrict the user to enter only upto 10 for the column 7
        
        .EditorDate = CDate("1/1/2000") 'Set Calender date
        .EditorTime = TimeValue("14:00") 'SEts editor time
        
        .Row = 1 'Setting the Row property to 1
        .Col = 1 'Setting the Col property to 1
        
        .ComboDropButtonStyle = fmDropButtonStyleEllipsis 'Sets the combo button type to ellipse
        
        .SetMaxLength(9) = 3 'Sets the maxlength of 9th column to 3
    End With

End Sub
Private Sub SetGrid2()
With GridEditor2
    .TextMatrix(0, 0) = "#"
    .TextMatrix(0, 1) = "User Name"
    .TextMatrix(0, 2) = "Read"
    .TextMatrix(0, 3) = "Write"
    .TextMatrix(0, 4) = "Full Access"
    For i = 2 To .Cols - 1
        .SetColObject(i) = CheckBox
    Next
    .SetColObject(1) = ComboBox
End With
End Sub
Private Sub Abt_Click()
    GridEditor1.ShowAbout
End Sub

Private Sub Command1_Click()
    Frm_Check_Box_Demo.Show
End Sub

Private Sub Form_Load()
Loading = True
Me.Top = 0: Me.Left = 0

Call SetGrid1
'Call SetGrid2
    Loading = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
With GridEditor1
    .Left = 0
    .Width = ScaleWidth
    .Height = ScaleHeight - .Top
End With
With Command1
    .Left = ScaleWidth - .Width - 10
End With
With Label1
    .Left = ScaleWidth / 2 - .Width / 2
    Label2.Left = .Left + 20
    Label2.Top = .Top + 30
End With
End Sub

Private Sub GridEditor1_AdvanceToNextCell(Cancel As Boolean)
'**************************************************************************************
'This event will trigger when the focus leaves from one cell to another cell inside the _
grid. So any validation for that specific cell can be done in this event.If you set _
Cancel = true then the user cannot move to next cell.

'This event will also trigger when the form loads. So we have to restrict to fire this event _
when the form loads. Loading boolean variable is used for that. It becomes true when form _
starts loading.
'**************************************************************************************

'If Loading Then Exit Sub
'
'    With GridEditor1
'        If .Col = 1 And Trim$(.TextMatrix(.Row, 1)) = "" Then
'            MsgBox "First column cannot left blank.", vbExclamation
'            Cancel = True
'        End If
'    End With
End Sub

Private Sub GridEditor1_Click()
    If GridEditor1.Col = 4 Then MsgBox "Editing is disabled for this cell.", vbExclamation
End Sub

Private Sub GridEditor1_ComboDropButtonClick()
    MsgBox "ComboButton dropDown clicked", vbInformation
End Sub

Private Sub GridEditor1_EditorChange()
    With GridEditor1
        If .Col = 2 And .EditorDate > Date Then
            MsgBox "Enter a date less than on equal to current date", vbExclamation
            .EditorDate = Date
            .Text = Date
        End If
    End With
End Sub



Private Sub GridEditor1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With Me
        Select Case GridEditor1.MouseCol
            Case 1
                .Caption = "Control is a Textbox; Col 1 cannot left blank"
            Case 2
                .Caption = "Control is a Datebox"
            Case 3
                .Caption = "Control is a Combobox"
            Case 4
                .Caption = "Cell 4 is Readonly."
            Case 5
                .Caption = "Control is a TimeBox"
            Case 6
                .Caption = "Control is a Textbox. Only numbers can be entered upto 1000.99"
            Case 7
                .Caption = "Control is a Textbox. Only numbers can be entered upto 10"
            Case Else
                .Caption = "Grid Editor Sample"
        End Select
    End With
    Timer1.Enabled = True
End Sub

Private Sub Opt1_Click(Index As Integer)
On Error Resume Next
With GridEditor1
    Select Case Index
        Case 0
            If .ShowEditor Then
                Opt1(0).Checked = False
                .ShowEditor = False
                .HideEditor
            Else
                Opt1(0).Checked = True
                .ShowEditor = True
            End If
        Case 1
            If .CreateNewRows Then
                Opt1(1).Checked = False
                .CreateNewRows = False
            Else
                Opt1(1).Checked = True
                .CreateNewRows = True
            End If
        Case 2
            .CalenderCustomFormat = InputBox("Enter Custom format : ", , "dd/MMM/yyyy")
            .CalenderFormat = dtpCustom
        Case 3
            .EditorDate = CDate(Format(InputBox("Enter any date : ", , Date), .CalenderCustomFormat))
        Case 4
            .EditorTime = TimeValue(Format(InputBox("Enter Time Value : ", , Time), "hh:mm:ss am/pm"))
        Case 6
            CommonDialog1.ShowColor
            .EditorBackColor = CommonDialog1.Color
        Case 7
            CommonDialog1.ShowColor
            .EditorForeColor = CommonDialog1.Color
        Case 8
            CommonDialog1.Flags = cdlCFBoth
            CommonDialog1.ShowFont
            Dim fnt As New StdFont
            fnt.Bold = CommonDialog1.FontBold
            fnt.Italic = CommonDialog1.FontItalic
            fnt.Size = CommonDialog1.FontSize
            fnt.Name = CommonDialog1.FontName
            fnt.Strikethrough = CommonDialog1.FontStrikethru
            fnt.Underline = CommonDialog1.FontUnderline
            Set .EditorFont = fnt
        Case 10
            .PrintView True
        Case 11
            .PrintGrid
    End Select
End With
End Sub

Private Sub Timer1_Timer()
Me.Caption = "Sample"
Timer1.Enabled = False
End Sub


