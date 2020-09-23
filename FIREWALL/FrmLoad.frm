VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLoad 
   ClientHeight    =   1005
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1005
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Progbar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Loading..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FrmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Greenlight As Boolean
Private Sub Form_Load()
Me.Visible = True
Progbar.Max = 10
Progbar.Value = 0
Greenlight = False

Progbar.Value = 5
loadit
Do Until Greenlight = True
DoEvents
Loop

Progbar.Value = 10

Form1.Visible = True
Unload Me

End Sub
Public Sub loadit()
DoEvents
Load Form1
Greenlight = True
End Sub
