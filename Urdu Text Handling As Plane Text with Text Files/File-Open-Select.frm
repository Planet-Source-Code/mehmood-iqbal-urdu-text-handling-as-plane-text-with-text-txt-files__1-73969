VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Dialog4 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5790
   Icon            =   "File-Open-Select.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Urdu_Text_Handling_As_Plane_Text.jcbutton jcbutton1 
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
      _extentx        =   3625
      _extenty        =   1085
      buttonstyle     =   13
      font            =   "File-Open-Select.frx":0A02
      backcolor       =   0
      caption         =   ""
      picturenormal   =   "File-Open-Select.frx":0A3A
      captioneffects  =   3
      picturealign    =   3
      tooltipbackcolor=   0
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5535
      Size            =   "9763;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5535
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "9763;1720"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Size            =   "9763;1085"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   315
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "Dialog4"
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

Private Sub Form_Load()

'Detect Text (*.Txt) files in App.Path & "\Urdu-TXT\" folder
'ans add them in Listbox
Files_Detection.ListFiles App.Path & "\Urdu-TXT\", "txt"

'Disable command button
If ListBox1.ListIndex = -1 Then jcbutton1.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

'If Form canceled with Right-side Close
Cancel = 1
Dialog4.Hide

End Sub

Private Sub jcbutton1_Click()

Dim File_Open_Path As String

'Set Path of selected file to be open
If ListBox1.ListIndex >= 0 Then

File_Open_Path = App.Path & "\Urdu-TXT\" & ListBox1.Object

End If

'Initialize File System Onjects (FSO)
Set Urdu_FSO = CreateObject("Scripting.FileSystemObject")

'Open file to be read
Set Urdu_Text_Stream = Urdu_FSO.OpenTextFile(File_Open_Path, ForReading, , TristateTrue)
      
      
'Read File Until EOF (End of File)
Do Until Urdu_Text_Stream.AtEndOfStream

      Urdu_Text = Urdu_Text + Urdu_Text_Stream.ReadLine
      
Loop

'{{

'Use Follwing Code Line If Line by Line Read Needed.

'Urdu_Text = Urdu_Text + Urdu_Text_Stream.ReadLine


'Add This Line (With Previous Line)
'as Needed to Read Number of Lines.

'Urdu_Text = Urdu_Text + vbCrLf + Urdu_Text_Stream.ReadLine

'}}

'Put Unicode file's Text in Unicode aware Textbox
Main_Form.TextBox1.Text = Urdu_Text

'Memory Clear
Set Urdu_FSO = Nothing
Set Urdu_Text_Stream = Nothing
Urdu_Text = ""

'Close Open dialog
Dialog4.Hide

End Sub

Private Sub ListBox1_Click()

'Enable command button, When listbox clicked
Dialog4.jcbutton1.Enabled = True

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
