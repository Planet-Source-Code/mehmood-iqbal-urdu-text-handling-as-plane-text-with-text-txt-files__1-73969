Attribute VB_Name = "Common_Functions"
'FSO Veriables

Dim Urdu_FSO As FileSystemObject

Dim Urdu_Text_Stream As TextStream

Dim Urdu_Text As String

'Diclarations for Fonts Installation
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_FONTCHANGE = &H1D
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Sub Unload_Forms()

'All form's unloading function for better & efficient closing

Unload Dialog5

'Detecting each Form & Unloading it
Dim ObjForm As Form
         
     For Each ObjForm In Forms
      
        Unload ObjForm
            
     Next

Set ObjForm = Nothing

'End All-Over Application
Close_Complete_Application

End Sub

Private Sub Close_Complete_Application()

'End All
End

End Sub
Public Sub Application_Intialize()

'Start Loading Progress
Dialog6.XP_ProgressBar1.Value = 5

'Using File system Objects (FSO) as Initializer
Set Urdu_FSO = CreateObject("Scripting.FileSystemObject")
Dialog6.XP_ProgressBar1.Value = 10

'Check necessory files to be exit
Files_Existance_Check
Dialog6.XP_ProgressBar1.Value = 20

'Opening file that contains Form's Caption with sequence
Set Urdu_Text_Stream = Urdu_FSO.OpenTextFile(App.Path & "\Urdu-Captions\Form_Captions.Dat", ForReading, , TristateTrue)
Dialog6.XP_ProgressBar1.Value = 30

'Setting all Form's Captions with sequential read
'CaptionW as Unicode Caption Function
Main_Form.CaptionW = Urdu_Text_Stream.ReadLine
Dialog.CaptionW = Urdu_Text_Stream.ReadLine
Dialog1.CaptionW = Urdu_Text_Stream.ReadLine
Dialog2.CaptionW = Urdu_Text_Stream.ReadLine
Dialog3.CaptionW = Urdu_Text_Stream.ReadLine
Dialog4.CaptionW = Urdu_Text_Stream.ReadLine
Dialog5.CaptionW = Urdu_Text_Stream.ReadLine
Dialog6.XP_ProgressBar1.Value = 40

'Diclaring new veriables for Control's Captions
Dim Top_Label_Caption As String
Dim Command1_Caption As String
Dim Command2_Caption As String
Dim Dialog_Label_Caption As String
Dim Dialog_Button_Caption As String
Dim Dialog1_Label_Caption As String
Dim Dialog1_Button_Caption As String
Dim Dialog2_Label_Caption As String
Dim Dialog2_Button_Caption As String
Dim Dialog3_Label_Caption As String
Dim Dialog3_Button_Caption As String
Dim Dialog4_Label_Caption As String
Dim Dialog4_Button_Caption As String
Dim Dialog5_Label_Caption As String
Dim Dialog5_Button1_Caption As String
Dim Dialog5_Button2_Caption As String
Dialog6.XP_ProgressBar1.Value = 50

'Opening file that contains Control's Captions
Set Urdu_Text_Stream = Urdu_FSO.OpenTextFile(App.Path & "\Urdu-Captions\Control_Captions.Dat", ForReading, , TristateTrue)
Dialog6.XP_ProgressBar1.Value = 60

'Putting Unicode (Urdu) Captions in diclared veriables
Top_Label_Caption = Top_Label_Caption + Urdu_Text_Stream.ReadLine
Command1_Caption = Command1_Caption + Urdu_Text_Stream.ReadLine
Command2_Caption = Command2_Caption + Urdu_Text_Stream.ReadLine
Dialog_Label_Caption = Dialog_Label + Urdu_Text_Stream.ReadLine
Dialog_Button_Caption = Dialog_Button_Caption + Urdu_Text_Stream.ReadLine
Dialog1_Label_Caption = Dialog1_Label_Caption + Urdu_Text_Stream.ReadLine
Dialog1_Button_Caption = Dialog1_Button_Caption + Urdu_Text_Stream.ReadLine
Dialog2_Label_Caption = Dialog2_Label_Caption + Urdu_Text_Stream.ReadLine
Dialog3_Label_Caption = Dialog3_Label_Caption + Urdu_Text_Stream.ReadLine
Dialog4_Label_Caption = Dialog4_Label_Caption + Urdu_Text_Stream.ReadLine
Dialog4_Button_Caption = Dialog4_Button_Caption + Urdu_Text_Stream.ReadLine
Dialog5_Label_Caption = Dialog5_Label_Caption + Urdu_Text_Stream.ReadLine
Dialog5_Button1_Caption = Dialog5_Button1_Caption + Urdu_Text_Stream.ReadLine
Dialog5_Button2_Caption = Dialog5_Button2_Caption + Urdu_Text_Stream.ReadLine
Dialog6.XP_ProgressBar1.Value = 70

'Setting Captions via veriables
Main_Form.Label1.Caption = Top_Label_Caption
Main_Form.Command1.Caption = Command1_Caption
Main_Form.Command2.Caption = Command2_Caption
Dialog.Label1.Caption = Dialog_Label_Caption
Dialog.jcbutton1.Caption = Dialog_Button_Caption
Dialog1.Label1.Caption = Dialog1_Label_Caption
Dialog1.jcbutton1.Caption = Dialog1_Button_Caption
Dialog2.Label1.Caption = Dialog2_Label_Caption
Dialog2.jcbutton1.Caption = Dialog1_Button_Caption
Dialog3.Label1.Caption = Dialog3_Label_Caption
Dialog3.jcbutton1.Caption = Dialog1_Button_Caption
Dialog4.Label1.Caption = Dialog4_Label_Caption
Dialog4.jcbutton1.Caption = Dialog4_Button_Caption
Dialog5.Label1.Caption = Dialog5_Label_Caption
Dialog5.jcbutton1.Caption = Dialog5_Button1_Caption
Dialog5.jcbutton2.Caption = Dialog5_Button2_Caption
Dialog6.XP_ProgressBar1.Value = 80

'Closing FSO Text Stream
Urdu_Text_Stream.Close

'Memory Clear
Set Urdu_FSO = Nothing
Set Urdu_Text_Stream = Nothing
Urdu_Text = ""
Dialog6.XP_ProgressBar1.Value = 90

'Clearing all diclared veriables
Top_Label_Caption = ""
Command1_Caption = ""
Command2_Caption = ""
Dialog_Label_Caption = ""
Dialog_Button_Caption = ""
Dialog1_Label_Caption = ""
Dialog1_Button_Caption = ""
Dialog2_Label_Caption = ""
Dialog2_Button_Caption = ""
Dialog3_Label_Caption = ""
Dialog3_Button_Caption = ""
Dialog4_Label_Caption = ""
Dialog4_Button_Caption = ""
Dialog5_Label_Caption = ""
Dialog5_Button1_Caption = ""
Dialog5_Button2_Caption = ""
Dialog6.XP_ProgressBar1.Value = 100

End Sub


Private Sub Files_Existance_Check()

'If files not exist not then install them form
'application folder with FSO.

'Checking Urdu font file in Fonts folder
If Urdu_FSO.FileExists("C:\WINDOWS\Fonts\Jameel Noori Nastaleeq.ttf") = False Then

'If Not Exit then install
Urdu_FSO.CopyFile (App.Path & "\Font\Jameel Noori Nastaleeq.ttf"), ("C:\WINDOWS\Fonts\"), True

End If

'Checking Urdu font file in Fonts folder
If Urdu_FSO.FileExists("C:\WINDOWS\Fonts\Besmellah 2.ttf") = False Then

'If Not Exit then install
Urdu_FSO.CopyFile (App.Path & "\Font\Besmellah 2.ttf"), ("C:\WINDOWS\Fonts\"), True

End If

'After Font's installation, its nessecory to tell all
'windows that font has been installed. So, calling GDI
'for that.

Dim res As Long
    
    'Add the font
    res = AddFontResource("C:\Windows\Fonts\Jameel Noori Nastaleeq.ttf")
    
    If res > 0 Then
    
        'Alert all windows that a font was added
        SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
      
     
    End If
    
    'Add the font
    res = AddFontResource("C:\Windows\Fonts\Besmellah 2.ttf")
    
    If res > 0 Then
   
       'Alert all windows that a font was added
       SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
     

    End If

'Put FM20.DLL file in App.Path & "\Files\" folder &
'Include following lines if you wants to install "FM20.DLL"
'for MS Form's 2.0 Controls, before Application Run. You can also
'put code to install necessory files before running of Application.

'Checking if Forms 2.0 Object Library file installed (comment)
'If Urdu_FSO.FileExists("C:\WINDOWS\system32\FM20.DLL") = False Then

'If Not Exit then install (comment)
'Urdu_FSO.CopyFile (App.Path & "\Files\FM20.DLL"), ("C:\WINDOWS\system32\"), True

'Else

'Do Nothing (comment)

'End If

End Sub
