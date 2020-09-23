Attribute VB_Name = "Files_Detection"
Public Sub ListFiles(strPath As String, Optional Extention As String)

'This Function will add & update *.Txt files in Listbox to open

    Dim File As String
    
    'check if there is a mistake of "\" in Path
    If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    'Extention is Optional, If not given then *.* is defualt
    If Trim$(Extention) = "" Then
    
        Extention = "*.*"
    
    'If Extension is given then use that
    ElseIf Left$(Extention, 2) <> "*." Then
    
        Extention = "*." & Extention
        
    End If
    
    'Detect Files
    File = Dir$(strPath & Extention)
    
    'Add all files (of given extention) in Listbox
    Do While Len(File)
    
        Dialog4.ListBox1.AddItem File
        
        File = Dir$
        
    Loop

'Update Path Label in Dialog4
Dialog4.Label2.Caption = strPath

'Clear Veriables
strPath = ""
Extention = ""

End Sub
