VERSION 5.00
Object = "*\AGridEditorControl.vbp"
Begin VB.Form FrmEntry 
   Caption         =   "Form2"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   6510
   WindowState     =   2  'Maximized
   Begin ActifOcx.GridEditor GridList 
      Height          =   3045
      Left            =   765
      TabIndex        =   0
      Top             =   420
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   5371
      ROWS            =   3
      COLS            =   3
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
      MOUSEICON       =   "FrmEntry.frx":0000
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
      ALLOWUSERRESIZING=   3
      ENABLED         =   -1  'True
      PrintOrient     =   0
   End
End
Attribute VB_Name = "FrmEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
On Error Resume Next
With GridList
    .ColWidth(0) = 400
    .Left = 0
    .Top = 0
    .Width = ScaleWidth
    .Height = ScaleHeight - .Top
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Closing = True
End Sub
