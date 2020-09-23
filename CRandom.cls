VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRandom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class Random
' MaRiØ Glez. Serrano. 10/01/02
' Clase para acceder a archivor de cualquier tamaño
' asi como a discos físicos directamente
' "\\.\A:" ->para abrir diskete
' si se accede directamente a Disco hay que leer
' de 512 en 512 (numero de bytes por sector)
' "COM1"..
' "LPT1"..

Public Enum W32F_Errors
  W32F_UNKNOWN_ERROR = 45600
  W32F_FILE_ALREADY_OPEN
  W32F_PROBLEM_OPENING_FILE
  W32F_FILE_ALREADY_CLOSED
  W32F_PROBLEM_SEEKING
End Enum

Private Const W32F_SOURCE = "Clase: Win32File Object"

Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2

Private Const GENERIC_WRITE          As Long = &H40000000
Private Const GENERIC_READ           As Long = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL  As Long = &H80
Private Const CREATE_ALWAYS          As Long = 2
Private Const OPEN_ALWAYS            As Long = 4
Private Const INVALID_HANDLE_VALUE   As Long = -1

Private Const FILE_BEGIN    As Long = 0
Private Const FILE_CURRENT  As Long = 1
Private Const FILE_END      As Long = 2

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Private Declare Function FormatMessage Lib "kernel32" _
        Alias "FormatMessageA" (ByVal dwFlags As Long, _
                                lpSource As Long, _
                                ByVal dwMessageId As Long, _
                                ByVal dwLanguageId As Long, _
                                ByVal lpBuffer As String, _
                                ByVal nSize As Long, _
                                Arguments As Any) As Long

Private Declare Function ReadFile Lib "kernel32" _
                        (ByVal hFile As Long, _
                         lpBuffer As Any, _
                         ByVal nNumberOfBytesToRead As Long, _
                         lpNumberOfBytesRead As Long, _
                         ByVal lpOverlapped As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
                        (ByVal hObject As Long) As Long

Private Declare Function WriteFile Lib "kernel32" _
                        (ByVal hFile As Long, _
                         lpBuffer As Any, _
                         ByVal nNumberOfBytesToWrite As Long, _
                         lpNumberOfBytesWritten As Long, _
                         ByVal lpOverlapped As Long) As Long

Private Declare Function CreateFile Lib "kernel32" _
        Alias "CreateFileA" (ByVal lpFileName As String, _
                             ByVal dwDesiredAccess As Long, _
                             ByVal dwShareMode As Long, _
                             ByVal lpSecurityAttributes As Long, _
                             ByVal dwCreationDisposition As Long, _
                             ByVal dwFlagsAndAttributes As Long, _
                             ByVal hTemplateFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" _
                        (ByVal hFile As Long, _
                         ByVal lDistanceToMove As Long, _
                         lpDistanceToMoveHigh As Long, _
                         ByVal dwMoveMethod As Long) As Long

Private Declare Function FlushFileBuffers Lib "kernel32" _
                        (ByVal hFile As Long) As Long

Private hFile As Long
Private sFName As String
Private fAutoFlush As Boolean
Private iPos As Long

Public Property Get FileHandle() As Long
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  'Else
   
  End If
   FileHandle = hFile
End Property

Public Property Get FileName() As String
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  FileName = sFName
End Property

Public Property Get IsOpen() As Boolean
  IsOpen = hFile <> INVALID_HANDLE_VALUE
End Property

Public Property Get AutoFlush() As Boolean
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  AutoFlush = fAutoFlush
End Property

Public Property Let AutoFlush(ByVal NewVal As Boolean)
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  fAutoFlush = NewVal
End Property

Public Sub OpenFile(ByVal sFileName As String)
  If hFile <> INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_OPEN, sFName
    Exit Sub
  End If
  hFile = CreateFile(sFileName, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ _
                Or FILE_SHARE_WRITE, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_PROBLEM_OPENING_FILE, sFileName
    Exit Sub
  End If
  sFName = sFileName
End Sub

Public Sub CloseFile()
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
    Exit Sub
  End If
  CloseHandle hFile
  sFName = ""
  fAutoFlush = False
  hFile = INVALID_HANDLE_VALUE
End Sub

Public Function ReadBytes(ByVal ByteCount As Long, Bytes() As Byte) As String
Dim BytesRead As Long ', bytes() As Byte
Dim asB As String

  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
    Exit Function
  End If
  ReDim Bytes(0 To ByteCount - 1) As Byte
  
  ReadFile hFile, Bytes(0), ByteCount, BytesRead, 0
  
  Dim i&
  For i = 0 To ByteCount - 1
   If Bytes(i) < 32 Then
      asB = asB & "·"
   Else
      asB = asB & Chr(Bytes(i))
   End If
  Next
  
  ReadBytes = asB
  
End Function

Public Sub WriteBytes(DataBytes() As Byte)
Dim fSuccess As Long, BytesToWrite As Long, BytesWritten As Long
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
    Exit Sub
  End If
  BytesToWrite = UBound(DataBytes) - LBound(DataBytes) + 1
  
  Me.SeekAbsolute 0, iPos
  fSuccess = WriteFile(hFile, DataBytes(LBound(DataBytes)), _
                       BytesToWrite, BytesWritten, 0)
  If fAutoFlush Then Flush
End Sub

Public Sub Flush()
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  Else
    FlushFileBuffers hFile
  End If
End Sub

Public Sub SeekAbsolute(ByVal HighPos As Long, ByVal LowPos As Long)
  If hFile = INVALID_HANDLE_VALUE Then
     RaiseError W32F_FILE_ALREADY_CLOSED
  Else
     LowPos = SetFilePointer(hFile, LowPos, HighPos, FILE_BEGIN)
  End If
  iPos = LowPos
End Sub

Public Sub SeekRelative(ByVal Offset As Long)
Dim TempLow As Long, TempErr As Long
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
    Exit Sub
  End If
  TempLow = SetFilePointer(hFile, Offset, ByVal 0&, FILE_CURRENT)
  If TempLow = -1 Then
    TempErr = Err.LastDllError
    If TempErr Then
      RaiseError W32F_PROBLEM_SEEKING, "Error " & TempErr & "." & _
                                       vbCrLf & CStr(TempErr)
    End If
  End If
End Sub

Private Sub Class_Initialize()
  hFile = INVALID_HANDLE_VALUE
End Sub

Private Sub Class_Terminate()
  If hFile <> INVALID_HANDLE_VALUE Then CloseHandle hFile
End Sub

Private Sub RaiseError(ByVal ErrorCode As W32F_Errors, _
                       Optional sExtra)
Dim Win32Err As Long, Win32Text As String
  Win32Err = Err.LastDllError
  If Win32Err Then
    Win32Text = vbCrLf & "Error " & Win32Err & vbCrLf & _
                DecodeAPIErrors(Win32Err)
  End If
  Select Case ErrorCode
    Case W32F_FILE_ALREADY_OPEN
      Err.Raise W32F_FILE_ALREADY_OPEN, W32F_SOURCE, _
          "El Archivo '" & sExtra & "' ya está abierto." & Win32Text
    Case W32F_PROBLEM_OPENING_FILE
      Err.Raise W32F_PROBLEM_OPENING_FILE, W32F_SOURCE, _
          "Error Abriendo '" & sExtra & "'." & Win32Text
    Case W32F_FILE_ALREADY_CLOSED
      Err.Raise W32F_FILE_ALREADY_CLOSED, W32F_SOURCE, _
          "El Archivo ya está cerrado."
    Case W32F_PROBLEM_SEEKING
      Err.Raise W32F_PROBLEM_SEEKING, W32F_SOURCE, _
          "Error de Búsqueda." & vbCrLf & sExtra
    Case Else
      Err.Raise W32F_UNKNOWN_ERROR, W32F_SOURCE, _
         "Error Desconocido." & Win32Text
  End Select
End Sub

Private Function DecodeAPIErrors(ByVal ErrorCode As Long) As String
Dim sMessage As String, MessageLength As Long
  sMessage = Space$(256)
  MessageLength = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, _
                                ErrorCode, 0&, sMessage, 256&, 0&)
  If MessageLength > 0 Then
    DecodeAPIErrors = Left(sMessage, MessageLength)
  Else
    DecodeAPIErrors = "Error Desconocido."
  End If
End Function




