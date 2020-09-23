VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Sample Application"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5430
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu gs 
      Caption         =   "GridEditor Sample"
   End
   Begin VB.Menu cs 
      Caption         =   "Picklist Demo"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cs_Click()
On Error Resume Next
FrmConnection.Show
FrmConnection.SetFocus
End Sub

Private Sub gs_Click()
Frm_Gridsample.Show
End Sub

