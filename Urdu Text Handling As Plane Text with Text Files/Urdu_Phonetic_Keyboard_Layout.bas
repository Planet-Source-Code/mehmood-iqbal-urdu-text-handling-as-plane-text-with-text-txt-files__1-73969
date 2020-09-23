Attribute VB_Name = "Urdu_Phonetic_Keyboard_Layout"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'© Copyright 2011, All Rights of this module Reserved '
'By Author of this project, Muhammad Mehmood Iqbal.   '
'                                                     '
'Author of this project give you an Open-Permission   '
'to use this module free of cost in your personal &   '
'professional projects, with following terms &        '
'conditions.                                          '
'Terms & Conditions :                                 '
'                                                     '
'1:This Module is provided on "as it is" basess, & no '
'  claim by any person will be acceptable later.      '
'2:When use this module at anywhere, DON'T REMOVE     '
'  COPYRIGHT TERMS & CONDITIONS & AUTHOR'S WRITTEN    '
'  COMMENTS, TERMS & CONDITIONS.                      '
'3:Author of this module DOES'NT PERMIT ANY PERSON TO '
'  PUBLISH THIS MODULE WITH ANY OTHER NAME & ON ANY   '
'  OTHER PLACE ELSE PSC.                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BY AUTHOR :                                           '
'Alhamdo-Lillah ! With the help of ALLAH, this common  '
'module has made by me. That is a very usefull         '
'achievement in the Urdu Programming series projects   '
'to make Urdu Programming Possible!. That was my       '
'try, comming from the start of Urdu Programming that  '
'i make a common module that completly define Urdu     '
'Keyboard layout, And it may common for (n) number of  '
'Textboxes. And now, Succeeded in that.                '
'Basically, this module designed for Unicode aware     '
'Textboxes that can recognize Urdu language characters.'
'Keydown & KeyPress avents are defined for Urdu        '
'character's recognization. I only prefer to use this  '
'module for Textboxes, because it is designed for that.'
'If you use for other controls, so there is no working '
'guaranty for them.                                    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'HOW TO USE Urdu_Phonetic_Keyboard_Layout:             '
'                                                      '
'There are some pre-requirement & conditions to use    '
'this module, as given in folowing :                   '
'1:>Textbox(s) that can be used, Must be Unicode Aware '
'   and can recognize Urdu language's characters.      '
'2:>Multiline Property of the Textbox(s) Must be True  '
'   before use, for better results.                    '
'3:>Jameel Noori Nastaleeq Urdu font is recomended for '
'   Textbox(s), for better & fine results.             '
'Ater these terms & conditions, you only put following '
'code lines as given in exaple code :                  '
'                                                      '
'KeyDown Event:                                        '''''''''''''''''''''''''''''
'Private Sub TextBox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)  '                     '
'Urdu_Phonetic_Keyboard.KeyDown TextBox1, KeyCode                                  '
'End Sub                                                                           '
'                                                                                  '
'KeyPress Event :                                                                  '
'Private Sub TextBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)                  '
'Urdu_Phonetic_Keyboard.KeyPress TextBox1, KeyAscii                                '
'End Sub                                               '''''''''''''''''''''''''''''                           '
'                                                      '
'Upper sample code is for MS Forms 2.0 Object Library  '
'Textbox(s) in VB6.Dont forget to define KeyCode &     '
'KeyAscii veriables before you use third part control. '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'There are some other Urdu Programming solutions are   '
'under-progress. And will be uploaded later on Planet  '
'Source Code.                                          '
'Any feedback by you will give me new ways to work,    '
'So, send me feedback about the Urdu Programming       '
'Project's series & specially about this module        '
'project. Waiting for your feedbacks on my email.      '
'Thank You.                                            '
'Regards,                                              '
'              Muhammd Mehmood Iqbal                   '
'               ME_IQ_TM@Yahoo.Com                     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'Veriables for whole module

'Module's private veriables for KeyDown event
Private KeyCode_Value As Integer
Private KeyDown_Value ' As String 'You can set

'Module's private veriables for KeyPress event
Private KeyAscii_Value As Integer
Private Unicode_Value ' As String 'You can set

Private Function Common_KeyDown()

'''''''''''''''''''''''''''''''''''''''''''''''
' There are the KeyDown-Event behaviours for  '
' Enter, Space, Tab & Delete keys to do a     '
' normal Behavior in Obj_Textbox, keys will   '
' behave as Normal Text writing behavior.     '
'                                             '
'            Muhammad Mehmood Iqbal           '
''             ME_IQ_TM@Yahoo.Com             '
'''''''''''''''''''''''''''''''''''''''''''''''

        'Space Key Behavior
        If KeyCode_Value = 32 Then
        Urdu_Phonetic_Keyboard_Layout.KeyDown_Value = "&H20"
      

        
        'Enter Key Behavior
        ElseIf KeyCode_Value = 13 Then
        Urdu_Phonetic_Keyboard_Layout.KeyDown_Value = "&HA"

        
        'Horizontal Tab Behavior
        ElseIf KeyCode_Value = 9 Then
        Urdu_Phonetic_Keyboard_Layout.KeyDown_Value = "&H9"

        
        'Delete Key Behavior
        ElseIf KeyCode_Value = 127 Then
        Urdu_Phonetic_Keyboard_Layout.KeyDown_Value = "&H7F"

        Else
        GoTo End_Fun:
        
        
        End If
        
'This Function Got End There
End_Fun:
End Function

Private Function Get_Unicode()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This Function will get veriable "KeyAscii_Value" value'
'from memory & will put-back a Unicode Hex value for a '
'specified Urdu character, as string to the veriable   '
'"Unicode_Value". That is the Main function of this    '
'Module.                                               '
'You can edit the existing Phoenetic Keyboard Layout as'
'you want.You can also make a new-one Keyboard Layout  '
'by changing Unicode values(Unicode_Value).            '
'                                                      '
'         Enjoy Using Urdu Keyboard Layout !           '
'                                                      '
'               Muhammad Mehmood Iqbal                 '
'                ME_IQ_TM@Yahoo.Com                    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        'For Small Letter's Behaviors

        'a Key Behavior
        If KeyAscii_Value = 97 Then
        
        Unicode_Value = "&H627"
        
        'b Key Behavior
        ElseIf KeyAscii_Value = 98 Then
        
        Unicode_Value = "&H628"
        
        'c Key Behavior
        ElseIf KeyAscii_Value = 99 Then
        
        Unicode_Value = "&H686"


        'd Key Behavior
        ElseIf KeyAscii_Value = 100 Then
        
        Unicode_Value = "&H62F"

        
        'e Key Behavior
        ElseIf KeyAscii_Value = 101 Then

        Unicode_Value = "&H639"

        
        'f Key Behavior
        ElseIf KeyAscii_Value = 102 Then

        Unicode_Value = "&H641"

        
        'g Key Behavior
        ElseIf KeyAscii_Value = 103 Then

        Unicode_Value = "&H6AF"

        
        'h Key Behavior
        ElseIf KeyAscii_Value = 104 Then

        Unicode_Value = "&H6BE"

        
        'i Key Behavior
        ElseIf KeyAscii_Value = 105 Then

        Unicode_Value = "&H6CC"

        
        'j Key Behavior
        ElseIf KeyAscii_Value = 106 Then

        Unicode_Value = "&H62C"

        
        'k Key Behavior
        ElseIf KeyAscii_Value = 107 Then

        Unicode_Value = "&H6A9"

        
        'l Key Behavior
        ElseIf KeyAscii_Value = 108 Then

        Unicode_Value = "&H644"

        
        'm Key Behavior
        ElseIf KeyAscii_Value = 109 Then

        Unicode_Value = "&H645"

        
        'n Key Behavior
        ElseIf KeyAscii_Value = 110 Then

        Unicode_Value = "&H646"
       
        
        'o Key Behavior
        ElseIf KeyAscii_Value = 111 Then

        Unicode_Value = "&H6C1"

        
        'p Key Behavior
        ElseIf KeyAscii_Value = 112 Then
        
        Unicode_Value = "&H67E"

        
        'q Key Behavior
        ElseIf KeyAscii_Value = 113 Then

        Unicode_Value = "&H642"

        
        'r Key Behavior
        ElseIf KeyAscii_Value = 114 Then
        
        Unicode_Value = "&H631"
        
        
        's Key Behavior
        ElseIf KeyAscii_Value = 115 Then
        
        Unicode_Value = "&H633"
        
        
        't Key Behavior
        ElseIf KeyAscii_Value = 116 Then
        
        Unicode_Value = "&H62A"
        
        
        'u Key Behavior
        ElseIf KeyAscii_Value = 117 Then
        
        Unicode_Value = "&H621"
        
        
        'v Key Behavior
        ElseIf KeyAscii_Value = 118 Then
        
        Unicode_Value = "&H637"
        
        
        'w Key Behavior
        ElseIf KeyAscii_Value = 119 Then
        
        Unicode_Value = "&H648"
        
        
        'x Key Behavior
        ElseIf KeyAscii_Value = 120 Then
        
        Unicode_Value = "&H634"
        
        
        'y Key Behavior
        ElseIf KeyAscii_Value = 121 Then
        
        Unicode_Value = "&H6D2"
        
        
        'z Key Behavior
        ElseIf KeyAscii_Value = 122 Then
        
        Unicode_Value = "&H632"
        
        
        
        ' For Capital Latter's Behaviors
        
        'A Key Behavior
        ElseIf KeyAscii_Value = 65 Then
        
        Unicode_Value = "&H622"
        
        
        'B Key Behavior
        ElseIf KeyAscii_Value = 66 Then
        
        Unicode_Value = "&HFBB0"
        
        
        'C Key Behavior
        ElseIf KeyAscii_Value = 67 Then
        
        Unicode_Value = "&H62B"
        
        
        'D Key Behavior
        ElseIf KeyAscii_Value = 68 Then
        
        Unicode_Value = "&H688"
        
        
        'E Key Behavior
        ElseIf KeyAscii_Value = 69 Then
        
        Unicode_Value = "&HE001"
        
        
        'F Key Behavior
        ElseIf KeyAscii_Value = 70 Then
        
        Unicode_Value = "&H652"
        
        
        'G Key Behavior
        ElseIf KeyAscii_Value = 71 Then
        
        Unicode_Value = "&H63A"
        
        
        'H Key Behavior
        ElseIf KeyAscii_Value = 72 Then
        
        Unicode_Value = "&H62D"
        
        
        'I Key Behavior
        ElseIf KeyAscii_Value = 73 Then
        
        Unicode_Value = "&H670"
        
        
        'J Key Behavior
        ElseIf KeyAscii_Value = 74 Then
        
        Unicode_Value = "&H636"
        
        
        'K Key Behavior
        ElseIf KeyAscii_Value = 75 Then
        
        Unicode_Value = "&H62E"
        
        
        'L Key Behavior
        ElseIf KeyAscii_Value = 76 Then
        
        Unicode_Value = "&HFEFB"
        
        
        'M Key Behavior
        ElseIf KeyAscii_Value = 77 Then
        
        Unicode_Value = "&H66B"
        
        
        'N Key Behavior
        ElseIf KeyAscii_Value = 78 Then
        
        Unicode_Value = "&H6BA"
        
        
        'O Key Behavior
        ElseIf KeyAscii_Value = 79 Then
        
        Unicode_Value = "&H6C3"
        
        
        'P Key Behavior
        ElseIf KeyAscii_Value = 80 Then
        
        Unicode_Value = "&H64F"
        
        
        'Q Key Behavior
        ElseIf KeyAscii_Value = 81 Then
        
        Unicode_Value = "&H626"
        
        
        'R Key Behavior
        ElseIf KeyAscii_Value = 82 Then
        
        Unicode_Value = "&H691"
        
        
        'S Key Behavior
        ElseIf KeyAscii_Value = 83 Then
        
        Unicode_Value = "&H635"
        
        
        'T Key Behavior
        ElseIf KeyAscii_Value = 84 Then
        
        Unicode_Value = "&H679"
        
        
        'U Key Behavior
        ElseIf KeyAscii_Value = 85 Then
        
        Unicode_Value = "&H626"
        
        
        'V Key Behavior
        ElseIf KeyAscii_Value = 86 Then
        
        Unicode_Value = "&H638"
        
        
        'W Key Behavior
        ElseIf KeyAscii_Value = 87 Then
        
        Unicode_Value = "&HFDFA"
        
        
        'Z Key Behavior
        ElseIf KeyAscii_Value = 88 Then
        
        Unicode_Value = "&H698"
        
        
        'Y Key Behavior
        ElseIf KeyAscii_Value = 89 Then
        
        Unicode_Value = "&H601"
        
        
        'Z Key Behavior
        ElseIf KeyAscii_Value = 90 Then
        
        Unicode_Value = "&H630"
        
        
        
        'For Numaric Key's Behaviors
        
        '0 Key Behavior
        ElseIf KeyAscii_Value = 48 Then
        
        Unicode_Value = "&H6F0"
        
        
        '1 Key Behavior
        ElseIf KeyAscii_Value = 49 Then
        
        Unicode_Value = "&H6F1"
        
        
        '2 Key Behavior
        ElseIf KeyAscii_Value = 50 Then
        
        Unicode_Value = "&H6F2"
        
        
        '3 Key Behavior
        ElseIf KeyAscii_Value = 51 Then
        
        Unicode_Value = "&H6F3"
        
        
        '4 Key Behavior
        ElseIf KeyAscii_Value = 52 Then
        
        Unicode_Value = "&H6F4"
        
        
        '5 Key Behavior
        ElseIf KeyAscii_Value = 53 Then
        
        Unicode_Value = "&H6F5"
        
        
        '6 Key Behavior
        ElseIf KeyAscii_Value = 54 Then
        
        Unicode_Value = "&H6F6"
        
        
        '7 Key Behavior
        ElseIf KeyAscii_Value = 55 Then
        
        Unicode_Value = "&H6F7"
        
        
        '8 Key Behavior
        ElseIf KeyAscii_Value = 56 Then
        
        Unicode_Value = "&H6F8"
        
        
        '9 Key Behavior
        ElseIf KeyAscii_Value = 57 Then
        
        Unicode_Value = "&H6F9"
        
        
        ' Numaric Keys with 'Shift' Behavior
        
        ') Key Behavior
        ElseIf KeyAscii_Value = 41 Then
        
        Unicode_Value = "&H29"
        
        
        '! Key Behavior
        ElseIf KeyAscii_Value = 33 Then
        
        Unicode_Value = "&H21"
        
        
        '@ Key Behavior
        ElseIf KeyAscii_Value = 64 Then
        
        Unicode_Value = "&H40"
        
        
        '# Key Behavior
        ElseIf KeyAscii_Value = 35 Then
        
        Unicode_Value = "&H23"
        
        
        '$ Key Behavior
        ElseIf KeyAscii_Value = 36 Then
        
        Unicode_Value = "&H24"
        
        
        '% Key Behavior
        ElseIf KeyAscii_Value = 37 Then
        
        Unicode_Value = "&H64A"
        
        
        '^ Key Behavior
        ElseIf KeyAscii_Value = 94 Then
        
        Unicode_Value = "&H5E"
        
        
        '& Key Behavior
        ElseIf KeyAscii_Value = 38 Then
        
        Unicode_Value = "&H26"
        
        
        '* Key Behavior
        ElseIf KeyAscii_Value = 42 Then
        
        Unicode_Value = "&H66D"
        
        
        '( Key Behavior
        ElseIf KeyAscii_Value = 40 Then
        
        Unicode_Value = "&H28"
        
        
        
        'For Special Characters
        
        'Symbols
        
        '? Key Behavior
        ElseIf KeyAscii_Value = 63 Then
        
        Unicode_Value = "&H61F"
        
        
        '/ Key Behavior
        ElseIf KeyAscii_Value = 47 Then
        
        Unicode_Value = "&H2F"
        
        
        ', Key Behavior
        ElseIf KeyAscii_Value = 44 Then
        
        Unicode_Value = "&H60C"
        
        
        '. Key Behavior
        ElseIf KeyAscii_Value = 46 Then
        
        Unicode_Value = "&H640"
        
        
        '_ Key Behavior
        ElseIf KeyAscii_Value = 95 Then
        
        Unicode_Value = "&H5F"
        
        
        '- Key Behavior
        ElseIf KeyAscii_Value = 45 Then
        
        Unicode_Value = "&H2D"
        
        
        '+ Key Behavior
        ElseIf KeyAscii_Value = 43 Then
        
        Unicode_Value = "&H2B"
        
        
        '= Key Behavior
        ElseIf KeyAscii_Value = 61 Then
        
        Unicode_Value = "&H3D"
        
        
        ': Key Behavior
        ElseIf KeyAscii_Value = 58 Then
        
        Unicode_Value = "&H3A"
        
        
        '; Key Behavior
        ElseIf KeyAscii_Value = 59 Then
        
        Unicode_Value = "&H201C"
        
        
        '< Key Behavior
        ElseIf KeyAscii_Value = 60 Then
        
        Unicode_Value = "&H64E"
        
        
        '> Key Behavior
        ElseIf KeyAscii_Value = 62 Then
        
        Unicode_Value = "&H650"
        
        
        '{ Key Behavior
        ElseIf KeyAscii_Value = 123 Then
        
        Unicode_Value = "&H2018"
        
        
        '} Key Behavior
        ElseIf KeyAscii_Value = 125 Then
        
        Unicode_Value = "&H2019"
        
        
        '[ Key Behavior
        ElseIf KeyAscii_Value = 91 Then
        
        Unicode_Value = "&H5B"
        
        
        '] Key Behavior
        ElseIf KeyAscii_Value = 93 Then
        
        Unicode_Value = "&H5D"
        
        
        '| Key Behavior
        ElseIf KeyAscii_Value = 124 Then
        
        Unicode_Value = "&H7C"
        
        
        '\ Key Behavior
        ElseIf KeyAscii_Value = 92 Then
        
        Unicode_Value = "&H5C"
        
        
        '~ Key Behavior
        ElseIf KeyAscii_Value = 126 Then
        
        Unicode_Value = "&H64B"
        
        
        '` Key Behavior
        ElseIf KeyAscii_Value = 96 Then
        
        Unicode_Value = "&H64D"
        
        
        '" Key Behavior
        ElseIf KeyAscii_Value = 34 Then
        
        Unicode_Value = "&H2190"
        
        
        '' Key Behavior
        ElseIf KeyAscii_Value = 39 Then
        
        Unicode_Value = "&H201D"
        
        
        End If
        
        KeyAscii_Value = 0
        
        'This Function Got End There

        
End Function
Public Sub KeyDown(Obj_Textbox As Object, KeyCode)

'Basic Public KeyDown Function that will be called by an
'external Object (Textbox etc), for KeyDown Event.

'Selec Case for KeyCode
Select Case KeyCode

'Selecting Case only for keys pressed Spacebar, Enter
'Tab-bar or Delete respectively.
Case 32, 13, 9, 127
      
      'Putting comming KeyCode value in veriable
      Urdu_Phonetic_Keyboard_Layout.KeyCode_Value = KeyCode
      
      'Getting Common Keydown behavior for these four keys
      Urdu_Phonetic_Keyboard_Layout.Common_KeyDown

      'Putting that behaviour in calling object (Textbox etc)
      Obj_Textbox = Obj_Textbox + ChrW(Urdu_Phonetic_Keyboard_Layout.KeyDown_Value)

      'Clearing veriables for next comming key
      KeyCode = 0
      Urdu_Phonetic_Keyboard_Layout.KeyDown_Value = ""
      
'Do Nothing if upper four keys are not pressed
Case Else

'End of Selecting Cases
End Select

End Sub

Public Function KeyPress(Obj_Textbox As Object, KeyAscii)

'Basic Public KeyPress Function that will be called by an
'external object (like Textbox etc), for Keypress Event.

'If Key is not 32 (Spacebar) then process
'Its because to avoide from some runtime errors.
If Not KeyAscii = 32 Then

    'Putting comming KeyAscii value in veriable
    Urdu_Phonetic_Keyboard_Layout.KeyAscii_Value = KeyAscii

    'Getting KeyAscii Unicode value
    Urdu_Phonetic_Keyboard_Layout.Get_Unicode

    'Putting Unicode Hex value (As Urdu Character) in calling
    'object (Textbox etc)
    Obj_Textbox.Text = Obj_Textbox.Text + ChrW(Urdu_Phonetic_Keyboard_Layout.Unicode_Value)

    'Clearing veriables for next comming key
    KeyAscii = 0
    Urdu_Phonetic_Keyboard_Layout.Unicode_Value = ""

End If

'End of the Function
End Function

'The Whole Module Gots End There
