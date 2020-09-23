VERSION 5.00
Begin VB.Form Frm_Grid_Preview 
   Caption         =   "Preview"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   Icon            =   "Frm_Grid_Preview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4110
      Top             =   225
   End
   Begin VB.PictureBox picScroll 
      Height          =   3375
      Left            =   60
      ScaleHeight     =   3315
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   810
      Width           =   4695
      Begin VB.VScrollBar vscScroll 
         Height          =   2535
         LargeChange     =   15
         Left            =   4290
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   255
      End
      Begin VB.HScrollBar hscScroll 
         Height          =   255
         LargeChange     =   15
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3090
         Width           =   4575
      End
      Begin VB.PictureBox picTarget 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   0
         ScaleHeight     =   2625
         ScaleWidth      =   3825
         TabIndex        =   3
         Top             =   0
         Width           =   3855
      End
   End
End
Attribute VB_Name = "Frm_Grid_Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cTP As clsTablePrint 'For Print Grid
Attribute cTP.VB_VarHelpID = -1
'The dimensions of the DIN A4 paper size in Twips:
Const A4Height = 16840, A4Width = 11907

'To get the scroll width:
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CYHSCROLL = 3
Private Const SM_CXVSCROLL = 2

Private Sub chkColWidth_Click()
Screen.MousePointer = vbHourglass
    Call InitializePictureBoxToPreview
    Set cTP = New clsTablePrint
    Call cmdRefresh_ClicktoPreview
    Frm_Grid_Preview.Show
Screen.MousePointer = vbDefault

End Sub

'**********************************************************************
'Functions used for Print Grid to a printer

Private Sub PrintGrid()
On Error Resume Next
Screen.MousePointer = vbHourglass
    'Call InitializePictureBox
    'Set cTP = New clsTablePrint
    'Call cmdRefresh_Click
'    Call cmdPrint_Click
'Screen.MousePointer = vbDefault
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
        
'        SetScrollBarsToPreview
    End With
End Sub

Private Sub cmdRefresh_ClicktoPreview()
    
    'Read the FlexGrid:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    With Frm_Grid_Preview
'    ImportFlexGrid cTP, EditorGrid, IIf((.chkColWidth.Value = vbChecked), .picTarget.ScaleWidth - 2 * 567, -1)
    
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
End Sub

Private Sub SetScrollBars()
    hscScroll.Max = (picTarget.Width - picScroll.ScaleWidth + vscScroll.Width) / 120 + 1
    vscScroll.Max = (picTarget.Height - picScroll.ScaleHeight + hscScroll.Height) / 120 + 1
End Sub
'*******************************End of functions used for PrintGrid***************************************

Private Sub Form_Load()
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
On Error Resume Next
With picScroll
    .Width = ScaleWidth
    .Top = 30
    .Left = 0
    .Height = ScaleHeight - 30
End With
With Frm_Grid_Preview
    .hscScroll.Max = (.picTarget.Width - .picScroll.ScaleWidth + .vscScroll.Width) / 120 + 1
    .vscScroll.Max = (.picTarget.Height - .picScroll.ScaleHeight + .hscScroll.Height) / 120 + 1
   
    .hscScroll.Left = 0
    .hscScroll.Width = Me.ScaleWidth - (vscScroll.Width + 40)
    .hscScroll.Top = Me.ScaleHeight - (.hscScroll.Height + picScroll.Top + 40)
    
    .vscScroll.Top = 0
    .vscScroll.Height = Me.ScaleHeight - (.hscScroll.Height + 40)
    .vscScroll.Left = Me.ScaleWidth - .vscScroll.Width - 50
    
End With
End Sub

Private Sub hscScroll_Change()
picTarget.Left = -hscScroll.Value * 120
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Frm_Grid_Preview.SetFocus
Timer1.Enabled = False
End Sub

Private Sub vscScroll_Change()
picTarget.Top = -CSng(vscScroll.Value) * 120
End Sub
