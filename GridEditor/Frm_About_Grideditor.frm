VERSION 5.00
Begin VB.Form Frm_About_Editor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About GridEditor"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   60
      Picture         =   "Frm_About_Grideditor.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   285
      TabIndex        =   4
      Top             =   60
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   360
      Left            =   4245
      TabIndex        =   2
      Top             =   2925
      Width           =   1005
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "shammiac@yahoo.com"
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
      Left            =   15
      TabIndex        =   7
      Top             =   1620
      Width           =   5205
   End
   Begin VB.Label Label5 
      Caption         =   "You are free to use this class in your   #"
      Height          =   1215
      Left            =   150
      TabIndex        =   6
      Top             =   2070
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Editor ActiveX Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   360
      TabIndex        =   5
      Top             =   60
      Width           =   2760
   End
   Begin VB.Line Line3 
      X1              =   -60
      X2              =   5670
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   -60
      X2              =   5685
      Y1              =   1950
      Y2              =   1965
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Freeware"
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
      Left            =   15
      TabIndex        =   3
      Top             =   1260
      Width           =   5205
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ActiveX Msflexgrid Editor Control"
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
      Left            =   15
      TabIndex        =   1
      Top             =   870
      Width           =   5205
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Left            =   15
      TabIndex        =   0
      Top             =   495
      Width           =   5205
   End
End
Attribute VB_Name = "Frm_About_Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub CheckBox1_Click()
'If CheckBox1.Value Then
'    SaveSetting "GridEditor", "Validate", "Show", 1
'Else
'    SaveSetting "GridEditor", "Validate", "Show", 0
'End If
'End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Label2.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Label5 = "You are free to use this control in your own projects . If you redistribute this control " & _
            " a notification would be  appreciated. This control comes with absolutely NO " & _
            " warranty ! Use it at your own risk !!! If there are any bugs regarding this control please mail to me. "
End Sub




