VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Dialog6 
   BorderStyle     =   0  'None
   ClientHeight    =   615
   ClientLeft      =   2715
   ClientTop       =   3315
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2640
      Top             =   1320
   End
   Begin Urdu_Text_Handling_As_Plane_Text.XP_ProgressBar XP_ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
   End
   Begin MSForms.Label Label1 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      Size            =   "6800;1085"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "Dialog6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Timer1_Timer()

'This will show loading progress bar before
'Main Form shows.

'Setting XP_Progressbar properties
Dialog6.XP_ProgressBar1.Min = 1
Dialog6.XP_ProgressBar1.Max = 100
Dialog6.XP_ProgressBar1.Value = 0

'Calling Application Initialize Function
'This Function will automatically set % loaded
'in progressbar.
Common_Functions.Application_Intialize

'When Loading completed,
If Dialog6.XP_ProgressBar1.Value = 100 Then

'Disable Timer for next time
Timer1.Enabled = False

'Hide loading progress-bar dialog
Dialog6.Hide

'Show Main Form
Main_Form.Show

End If

End Sub
