VERSION 5.00
Object = "*\AGridEditorControl.vbp"
Begin VB.Form Frm_Check_Box_Demo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Check Box Demo"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin ActifOcx.GridEditor GridEditor2 
      Height          =   2145
      Left            =   75
      TabIndex        =   2
      Top             =   285
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   3784
      ROWS            =   5
      COLS            =   5
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
      MOUSEICON       =   "Frm_Check_Box_Demo.frx":0000
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
      WORDWRAP        =   0   'False
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
   Begin VB.CheckBox Check1 
      Caption         =   "Create New Rows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3645
      TabIndex        =   1
      Top             =   0
      Width           =   1965
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Demonstrating Check Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3105
   End
End
Attribute VB_Name = "Frm_Check_Box_Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        GridEditor2.CreateNewRows = True
    Else
        GridEditor2.CreateNewRows = False
    End If
End Sub

Private Sub Form_Load()
With GridEditor2
    .CreateNewRows = False 'Will not create new rows
    .TextMatrix(0, 0) = "#"
    .TextMatrix(0, 1) = "User Name"
    .TextMatrix(0, 2) = "Read"
    .TextMatrix(0, 3) = "Write"
    .TextMatrix(0, 4) = "Full Access"
    For i = 2 To .Cols - 1
        .SetColObject(i) = CheckBox
    Next
    .SetColObject(1) = ComboBox
    .RowHeight(0) = 500
        .Row = 0
        For i = 1 To .Cols - 1
            .Col = i
            .CellAlignment = 4 'Set Alignment to Center - Center
            .CellFontBold = True
        Next
        .Row = 1
        .Col = 1
        .ColWidth(0) = 300
        .ColWidth(1) = 2000
        .GenerateGridNumber
        .ComboAdditem "Jacobs, Russell"
        .ComboAdditem "Metzger, Philip W."
        .ComboAdditem "Gardner, Juanita Mercado"
        .ComboAdditem "Gabriel, Richard P."
        If GridEditor2.CreateNewRows Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
End With
End Sub

Private Sub GridEditor2_Click()
'*****************************************************************
'Date        Developer           Comments                        *
'20/9/01     SHAHIM.A.C       Initial creation                   *
'*****************************************************************

With GridEditor2
    If .Row = 0 Then Exit Sub
    Select Case .Col
        Case 2
            If Val(.TextMatrix(.Row, 2)) = 1 Then
                .TextMatrix(.Row, 3) = 0: .TextMatrix(.Row, 4) = 0
            End If
        Case 3
            If Val(.TextMatrix(.Row, 3)) = 1 Then
                .TextMatrix(.Row, 2) = 0: .TextMatrix(.Row, 4) = 0
            End If
        Case 4
            If Val(.TextMatrix(.Row, 4)) = 1 Then
                .TextMatrix(.Row, 2) = 0: .TextMatrix(.Row, 3) = 0
            End If
    End Select
    .GenerateGridNumber
End With
End Sub

