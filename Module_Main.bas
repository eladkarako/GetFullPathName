Attribute VB_Name = "Module_Main"
Option Explicit

Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Function GetFullPath(strFileName As String) As String
  'KPD-Team 1999
  'URL: http://www.allapi.net/
  'E-Mail: KPDTeam@Allapi.net
  Dim Buffer As String, Ret As Long
  
  On Error Resume Next
  GetFullPath = ""

  'create a buffer
  Buffer = Space$(255)
  'copy the current directory to the buffer and append 'myfile.ext'
  Ret = GetFullPathName(strFileName, 255, Buffer, "")
  'remove the unnecessary chr$(0)'s
  Buffer = Left$(Buffer, Ret)
  'show the result
  GetFullPath = Buffer
End Function

Public Sub Main()
  On Error Resume Next

  Dim s As String
  s = ""
  s = Command$
  s = Replace$(s, Chr$(34), "")

  s = Left$(s, InStr(vbNull, s, vbNullChar, vbBinaryCompare) - vbNull)
  s = GetFullPath(s)

  WriteStdOut s
End Sub

