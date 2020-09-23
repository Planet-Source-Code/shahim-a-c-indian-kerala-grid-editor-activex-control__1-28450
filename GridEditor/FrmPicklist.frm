VERSION 5.00
Object = "*\AGridEditorControl.vbp"
Begin VB.Form FrmPicklist 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   5280
   WindowState     =   2  'Maximized
   Begin ActifOcx.PickList PickList1 
      Height          =   4020
      Left            =   45
      TabIndex        =   9
      Top             =   60
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   7091
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2610
      Left            =   3705
      TabIndex        =   1
      Top             =   300
      Width           =   1605
      Begin VB.OptionButton Option2 
         Caption         =   "Small Icon"
         Height          =   210
         Left            =   0
         TabIndex        =   0
         Top             =   942
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Large Icon"
         Height          =   240
         Left            =   0
         TabIndex        =   8
         Top             =   608
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Multi Select"
         Height          =   210
         Left            =   0
         TabIndex        =   7
         Top             =   2160
         Width           =   1305
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check Boxes"
         Height          =   210
         Left            =   0
         TabIndex        =   6
         Top             =   1854
         Width           =   1275
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Movable"
         Height          =   210
         Left            =   0
         TabIndex        =   5
         Top             =   1550
         Width           =   1425
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Show Text"
         Height          =   210
         Left            =   0
         TabIndex        =   4
         Top             =   1246
         Value           =   1  'Checked
         Width           =   1260
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Tools"
         Height          =   210
         Left            =   0
         TabIndex        =   3
         Top             =   304
         Value           =   1  'Checked
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Caption"
         Height          =   210
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1380
      End
   End
End
Attribute VB_Name = "FrmPicklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
PickList1.ShowCaption = Not PickList1.ShowCaption
End Sub

Private Sub Check2_Click()
PickList1.ShowTools = Not PickList1.ShowTools
If Check2.Value = 0 Then
    Option1.Enabled = False
    Option2.Enabled = False
Else
    Option1.Enabled = True
    Option2.Enabled = True
End If
End Sub

Private Sub Check3_Click()
PickList1.ShowText = Not PickList1.ShowText
End Sub

Private Sub Check4_Click()
PickList1.Movable = Not PickList1.Movable
End Sub

Private Sub Check5_Click()
PickList1.Checkboxes = Not PickList1.Checkboxes
End Sub

Private Sub Check6_Click()
PickList1.MultiSelect = Not PickList1.MultiSelect
End Sub

Private Sub Form_Load()
PickList1.ShowCaption = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
With Frame1
    .Left = ScaleWidth - .Width
End With
With PickList1
.Left = 0
.Top = 0
.Width = ScaleWidth - Frame1.Width - 50
.Height = ScaleHeight
End With
End Sub

Private Sub Option1_Click()
PickList1.IconType = LargeIcons
End Sub

Private Sub Option2_Click()
PickList1.IconType = SmallIcons
End Sub

Private Sub PickList1_OnCancelClicked()
Unload Me
End Sub

Private Sub PickList1_OnQueryClicked(Cancel As Boolean)
MsgBox "Query Clicked.", vbInformation
If MsgBox("Do you want to cancel Querying?", vbQuestion + vbYesNo) = vbYes Then
    Cancel = True
    MsgBox "Querying Cancelled.", vbInformation
End If
End Sub

Private Sub PickList1_OnRefreshClicked(Cancel As Boolean)
MsgBox "Refreshed Clicked.", vbInformation
If MsgBox("Do you want to cancel refreshing?", vbQuestion + vbYesNo) = vbYes Then
    Cancel = True
    MsgBox "Refreshing Cancelled.", vbInformation
End If
End Sub

Private Sub PickList1_OnSelectClicked()
Dim i&, j&, Msg$
With PickList1
If .SelectedCount = 0 Then
    MsgBox "Nothing selected.", vbExclamation
    Exit Sub
End If
If .SelectedCount > 1 Then
    For i = 1 To .RecordCount
        If .Selected(i) Then
            For j = 1 To .ColumnCount
                Msg = Msg & .TextMatrix(i, j) & Space(5)
            Next
            Msg = Msg & vbCr
        End If
    Next
Else
    For i = 1 To .ColumnCount
        Msg = Msg & "Selected Column No." & i & " : " & PickList1.SelectedItem(i) & vbCr
    Next
End If
MsgBox Msg
End With
End Sub
