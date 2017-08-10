Attribute VB_Name = "Module_Main"
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

Public Sub Main()
    'KPD-Team 2000
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim Buffer As String, Ret As Long
    'create a buffer
    Buffer = Space$(255)
    'copy the current directory to the buffer and append 'myfile.ext'
    Ret = GetFullPathName(Command, 255, Buffer, "")
    'remove the unnecessary chr$(0)'s
    Buffer = Left$(Buffer, Ret)
    'show the result
    WriteStdOut Buffer
End Sub

