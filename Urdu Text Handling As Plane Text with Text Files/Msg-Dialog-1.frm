VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Dialog1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2055
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6405
   Icon            =   "Msg-Dialog-1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Urdu_Text_Handling_As_Plane_Text.jcbutton jcbutton1 
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jameel Noori Nastaleeq"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   ""
      PictureNormal   =   "Msg-Dialog-1.frx":0A02
      PictureAlign    =   3
      CaptionEffects  =   3
      TooltipBackColor=   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5400
      Picture         =   "Msg-Dialog-1.frx":1414
      Top             =   120
      Width           =   480
   End
   Begin MSForms.Label Label2 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   6255
      Size            =   "11033;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Size            =   "10186;1085"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   315
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Function Declarations for CaptionW Property
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long) As Long
Private Const GWL_WNDPROC = -4
Private m_Caption As String

Private Sub Form_Unload(Cancel As Integer)

'If Form canceled with Right-side Close
Cancel = 1
Dialog1.Hide

End Sub

Private Sub jcbutton1_Click()

'File Saved Sucessfully then hide form & goto Main_Form
Dialog1.Hide
Main_Form.TextBox1.SetFocus

End Sub

Public Property Get CaptionW() As String

''''''''''''''''''''''''''''''''''''''''''''''''''''
'These code lines are By Vesa Piittinen aka Merri, '
'<vesa@piittinen.name> shared on VBForums & PSC.   '
''''''''''''''''''''''''''''''''''''''''''''''''''''

    CaptionW = m_Caption
    
End Property

Public Property Let CaptionW(ByRef NewValue As String)

''''''''''''''''''''''''''''''''''''''''''''''''''''
'These code lines are By Vesa Piittinen aka Merri, '
'<vesa@piittinen.name> shared on VBForums & PSC.   '
''''''''''''''''''''''''''''''''''''''''''''''''''''


    Static WndProc As Long, VBWndProc As Long
    m_Caption = NewValue
    
    ' Get window procedures if we don't have them
    If WndProc = 0 Then
    
        ' The default Unicode window procedure
        WndProc = GetProcAddress(GetModuleHandleW(StrPtr("user32")), "DefWindowProcW")
        
        ' Window procedure of this form
        VBWndProc = GetWindowLongA(hwnd, GWL_WNDPROC)
    End If
    
    ' Ensure we got them
    If WndProc <> 0 Then
    
        ' Replace form's window procedure with the default Unicode one
        SetWindowLongW hwnd, GWL_WNDPROC, WndProc
        
        ' Change form's caption
        SetWindowTextW hwnd, StrPtr(m_Caption)
        
        ' Restore the original window procedure
        SetWindowLongA hwnd, GWL_WNDPROC, VBWndProc
    Else
    
        ' No Unicode for us
        Caption = m_Caption
        
    End If
    
End Property
