Type=Exe
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\system32\stdole2.tlb#OLE Automation
Reference=*\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\WINDOWS\system32\scrrun.dll#Microsoft Scripting Runtime
Form=FSO.frm
Object={0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0; FM20.DLL
UserControl=Urdu-Button.ctl
Form=Msg-Dialog.frm
Form=Msg-Dialog-1.frm
Form=Msg-Dialog-2.frm
Form=Msg-Dialog-3.frm
Form=File-Open-Select.frm
Module=Files_Detection; Files_Detection.bas
Form=Msg-Dialog-4.frm
Module=Common_Functions; Common_Functions.bas
Startup="Form1"
Command32=""
Name="Urdu_Text_Handling_As_Plane_Text"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
VersionCompanyName="Sab Thinker & Active Org."
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1

[MS Transaction Server]
AutoRefresh=1
                                                                                                                                                                                                                                                                                                                                                                                                                                                                     
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
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   3
      TooltipBackColor=   0
   End
   Begin MSForms.Label Label1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
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
Attribute VB_Name = "Dialog5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub jcbutton1_Click()

Common_Functions.Unload_Forms

End Sub

Private Sub jcbutton2_Click()

Dialog5.Hide

End Sub
