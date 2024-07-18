Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants and functions used by this program.
Private Declare Function GetComputerNameA Lib "Kernel32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTempFileNameA Lib "Kernel32.dll" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTemporaryFileName As String) As Long
Private Declare Function GetTempPathA Lib "Kernel32.dll" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetUserNameA Lib "Advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long


'The constants used by this program.
Private Const MAX_COMPUTERNAME_LENGTH As Long = 32   'Defines the maximum allowed length for computer names.
Private Const MAX_PATH As Long = 256                 'Defines the maximum allowed length for paths.
Private Const MAX_STRING As Long = 65535             'Defines the maximum allowed length for strings.


'This procedure retrieves and returns the computer's name.
Private Function GetComputerName() As String
On Error GoTo ErrorTrap
Dim ComputerName As String
Dim ReturnValue As Long

   ComputerName = String$(MAX_COMPUTERNAME_LENGTH, vbNullChar)
   ReturnValue = GetComputerNameA(ComputerName, Len(ComputerName))

   If InStr(ComputerName, vbNullChar) > 0 Then
      ComputerName = Left$(ComputerName, InStr(ComputerName, vbNullChar) - 1)
   Else
      ComputerName = vbNullString
   End If

EndProcedure:
   GetComputerName = ComputerName
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function
'This procedure generates a temporary file name using the specified parameters and returns the result.
Private Function GetTemporaryFileName(TemporaryPath As String, Prefix As String, Optional UniqueNumber As Long = 0) As String
On Error GoTo ErrorTrap
Dim TemporaryFile As String
Dim UniqueNumberReturned As Long

   TemporaryFile = String$(MAX_PATH, vbNullChar)
   UniqueNumberReturned = GetTempFileNameA(TemporaryPath, Prefix, UniqueNumber, TemporaryFile)
   If InStr(TemporaryFile, vbNullChar) Then TemporaryFile = Left$(TemporaryFile, InStr(TemporaryFile, vbNullChar) - 1)
  
EndProcedure:
   GetTemporaryFileName = Prefix & Hex$(UniqueNumberReturned)
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function

'This procedure retrieves and returns the temporary file folder's path.
Private Function GetTemporaryPath() As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim TemporaryPath As String

   TemporaryPath = String$(MAX_PATH, vbNullChar)
   Length = GetTempPathA(Len(TemporaryPath), TemporaryPath)
   If Length > 0 Then TemporaryPath = Left$(TemporaryPath, Length)

EndProcedure:
   GetTemporaryPath = TemporaryPath
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function


'This procedure retrieves and returns the current user's name.
Private Function GetUserName() As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim UserName As String

   UserName = String$(MAX_STRING, vbNullChar)
   Length = GetUserNameA(UserName, Len(UserName))
   UserName = Left$(UserName, InStr(UserName, vbNullChar) - 1)
   
EndProcedure:
   GetUserName = UserName
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function



'This procedure handles any errors that occur.
Private Function HandleError(Optional ReturnPreviousChoice As Boolean = False) As Long
Dim Description As String
Dim ErrorCode As Long
Static Choice As Long

   Description = Err.Description
   ErrorCode = Err.Number
   On Error Resume Next
   If Not ReturnPreviousChoice Then
      Choice = MsgBox(Description & "." & vbCr & "Error code: " & CStr(ErrorCode), vbAbortRetryIgnore Or vbDefaultButton2 Or vbExclamation)
   End If
   
   If Choice = vbAbort Then End
   
   HandleError = Choice
End Function
'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
Dim Information As String
Dim TemporaryPath As String
   
   TemporaryPath = GetTemporaryPath()
   
   Information = "Computer name: " & GetComputerName() & vbCr
   Information = Information & "Current user: " & GetUserName() & vbCr
   Information = Information & "Temporary file name: " & GetTemporaryFileName(TemporaryPath, "tmp") & vbCr
   Information = Information & "Temporary file path: " & TemporaryPath & vbCr
   
   MsgBox Information, vbInformation
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub

