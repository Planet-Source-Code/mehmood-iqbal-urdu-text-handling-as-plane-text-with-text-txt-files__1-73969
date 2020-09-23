VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2190
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5790
   Icon            =   "Msg-Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleMode       =   0  'User
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Urdu_Text_Handling_As_Plane_Text.jcbutton jcbutton1 
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
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
      PictureNormal   =   "Msg-Dialog.frx":0A02
      PictureAlign    =   3
      CaptionEffects  =   3
      TooltipBackColor=   0
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5535
      VariousPropertyBits=   746604571
      MaxLength       =   20
      Size            =   "9763;873"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   615
      Left            =   0
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
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FSO Veriables
Dim Urdu_FSO As FileSystemObject

Dim Urdu_Text_Stream As TextStream

Dim Urdu_Text As String

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

Private Sub Form_Activate()

'Set focus to Textbox
Dialog.TextBox1.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)

'If Form canceled with Right-side Close
Cancel = 1
Dialog.Hide

End Sub

Private Sub jcbutton1_Click()

Dim Trim_Text As String

'Trim Right & Left of Textbox1.Text
Trim_Text = Trim$(TextBox1.Text)

'Put back a new Trimed-Text
TextBox1.Text = Trim_Text

Dim File_Path As String

'Specifying a path where file will be saved
File_Path = App.Path & "\" & "Urdu-TXT\" & TextBox1.Text & ".Txt"

'Initialize File System Objects (FSO)
Set Urdu_FSO = CreateObject("Scripting.FileSystemObject")

'Check if File already exist
If Urdu_FSO.FileExists(File_Path) Then

'Warning message
Dialog2.Show vbModal

'Check if File name is not given
ElseIf Dialog.TextBox1.Text = "" Then

'Warning message
Dialog3.Show vbModal

Else

'Creat a Unicode File, that will contain Unicode Data
'At last of the statment "True" tells FSO to creat a
'Unicode Text File
Set Urdu_Text_Stream = Urdu_FSO.CreateTextFile(File_Path, True, True)

'Change Mouse Pointer to HourGlass
Dialog.MousePointer = "11"

'Putting Data to Created Unicode File
Urdu_Text_Stream.WriteLine (Main_Form.TextBox1.Text)

'Close file
Urdu_Text_Stream.Close

'Clear FSO veriables
Set Urdu_FSO = Nothing
Set Urdu_Text_Stream = Nothing

'Close Save dialog, clear its Textbox
Dialog.Hide
Dialog.TextBox1.Text = ""

'Show Sucess information dialog
Dialog1.Show
Dialog1.Label2.Caption = File_Path

'Clear existing items in 'open file' dialog
'Updating it for new saved file.
Dialog4.ListBox1.Clear
Files_Detection.ListFiles App.Path & "\Urdu-TXT\", "txt"

'Clear veriable
File_Path = ""

'Set MousePointer to defualt
Dialog.MousePointer = "0"

End If

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
