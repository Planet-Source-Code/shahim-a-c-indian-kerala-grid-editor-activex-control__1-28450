VERSION 5.00
Object = "{1C21A313-B19D-11D5-90FF-0050BA341114}#1.11#0"; "ActifControls.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin ActifOcx.PickList PickList1 
      Height          =   6630
      Left            =   60
      TabIndex        =   2
      Top             =   75
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   11695
      ConnectionString1=   ""
      CHECKBOX        =   0   'False
      MULTISELECT     =   0   'False
      MainQuery       =   ""
      BeginProperty FONT12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FORECOLOR       =   -2147483640
      ShowCaption1    =   -1  'True
      LBLCAPTION      =   ""
      CAPBACKCOLOR    =   -2147483633
      CAPFORECOLOR    =   -2147483634
      BORDERCOLOR1    =   -2147483640
      BORDERCOLOR2    =   -2147483643
      CAPALIGN        =   0
      BeginProperty CAPFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MovablePicklist =   0   'False
      LblBackground   =   8388608
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   8265
      TabIndex        =   1
      Top             =   960
      Width           =   810
   End
   Begin ActifOcx.GridEditor GridEditor1 
      Height          =   2790
      Left            =   9300
      TabIndex        =   0
      Top             =   1770
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   4921
      ROWS            =   5
      COLS            =   6
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
      MOUSEICON       =   "Form1.frx":0000
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
      CREATENEWROWS   =   0   'False
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With GridEditor1
    .SaveGrid "c:\test.txt"
End With
End Sub

Private Sub Form_Load()
Dim i As New ADODB.Connection

i.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Works\Actif (RKG)\database(rkg)\Copy of school.mdb;Persist Security Info=False;Jet OLEDB:Database Password=svss"
PickList1.PicklistConnection = i
PickList1.RetrieveData "select * from student_mst"
End Sub

Private Sub PickList1_GotFocus()

End Sub

Private Sub PickList1_ItemCheck(ByVal Item As MSComctlLib.ListItem)

End Sub

Private Sub PickList1_ItemClick(ByVal Item As MSComctlLib.ListItem)

End Sub

Private Sub PickList1_OnQueryClicked(Cancel As Boolean)

End Sub

Private Sub PickList1_OnSelectClicked()

End Sub
