Attribute VB_Name = "modINI"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Global AcademyName As String
Global DataPath As String
Global DataFile As String
Global DsnFile As String
Global INIFile As String
'***************************************
'     ****************
'* Procedure Name: sReadINI*
'*======================================
'     ===============*
'*Returns a string from an INI file. To
'     use, call the *
'*functions and pass it the Section, Key
'     Name and INI*
'*File Name, [sRet=sReadINI(Section,Key1
'     ,INIFile)].*
'*val command. *
'***************************************
'     ****************

Function ReadINI(rSection, rKeyName As String, rFilename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(rSection, rKeyName, "", sRet, Len(sRet), rFilename))
End Function

'***************************************
'     ****************
'* Procedure Name: WriteINI*
'*======================================
'     ===============*
'*Writes a string to an INI file. To use
'     , call the *
'*function and pass it the sSection, sKe
'     yName, the New *
'*String and the INI File Name,*
'*[Ret=WriteINI(Section,Key,String,INIFi
'     le)]. *
'*Returns a 1 if there were no errors an
'     d *
'*a 0 if there were errors.*
'***************************************
'     ****************

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function
