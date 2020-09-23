VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Main_Form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Urdu Text Handling As Plane Text with Text (*.Txt) Files"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8610
   Icon            =   "Main Urdu Text Handling.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Main Urdu Text Handling.frx":0A02
   ScaleHeight     =   6000
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin Urdu_Text_Handling_As_Plane_Text.jcbutton Command2 
      Height          =   615
      Left            =   4560
      TabIndex        =   3
      Top             =   5160
      Width           =   2295
      _ExtentX        =   4048
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
      BackColor       =   15199212
      Caption         =   ""
      PictureNormal   =   "Main Urdu Text Handling.frx":1404
      PictureAlign    =   3
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Urdu_Text_Handling_As_Plane_Text.jcbutton Command1 
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   5160
      Width           =   2295
      _ExtentX        =   4048
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
      BackColor       =   15199212
      Caption         =   ""
      PictureNormal   =   "Main Urdu Text Handling.frx":1E16
      PictureAlign    =   3
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      TooltipBackColor=   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   4320
      X2              =   4320
      Y1              =   5760
      Y2              =   5160
   End
   Begin MSForms.Label Label1 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   -240
      Width           =   8655
      ForeColor       =   12582912
      Size            =   "15266;1931"
      FontName        =   "Besmellah 2"
      FontHeight      =   1440
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   8415
      VariousPropertyBits=   -1400879077
      Size            =   "14843;7011"
      SpecialEffect   =   3
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   480
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
End
Attribute VB_Name = "Main_Form"
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

Private Sub Command1_Click()

Dialog.Show vbModal

End Sub

Private Sub Command2_Click()

Dialog4.Show vbModal

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dialog5.Show vbModal

End Sub

Private Sub TextBox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

Urdu_Phonetic_Keyboard_Layout.KeyDown TextBox1, KeyCode

End Sub

Private Sub TextBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)

Urdu_Phonetic_Keyboard_Layout.KeyPress TextBox1, KeyAscii

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



