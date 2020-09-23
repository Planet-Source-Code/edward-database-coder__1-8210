Attribute VB_Name = "modAddIn"
' ==============================================================
' Module:       modAddIn
' Purpose:      Single procedure to add project to vbaddin.ini
' Execute:      From immediate window 'AddToINI'
' ==============================================================

Option Explicit
Declare Function WritePrivateProfileString& Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

'====================================================================
'this sub should be executed from the Immediate window
'in order to get this app added to the VBADDIN.INI file
'====================================================================
Sub AddToINI()
    Dim ErrCode As Long

    ErrCode = WritePrivateProfileString("Add-Ins32", "DatabaseCoder.Connect", "0", "vbaddin.ini")
    MsgBox "qbd DatabaseCoder has been added to the Add-Ins32.ini file" & vbCrLf _
    & "You should now select 'Make dbCoderAddIn.dll from the File menu."
    
End Sub

