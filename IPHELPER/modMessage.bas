Attribute VB_Name = "modMessage"
Option Explicit

Private Const LANG_NEUTRAL = &H0
Private Const SUBLANG_DEFAULT = &H1
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Function GetDLLMessage(ErrorNo As Long) As String
   Dim Flags As Long, Puffer As String
   Dim RetVal As Long, Sprache As Long
   Dim Fehler As Long
 
   Flags = FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS
   
   Sprache = LANG_NEUTRAL Or (SUBLANG_DEFAULT * 1024)
   
   Puffer = Space(256)
   Fehler = ErrorNo
 
   RetVal = FormatMessage(Flags, 0&, Fehler, Sprache, Puffer, Len(Puffer), 0&)
 
   If RetVal <> 0 Then GetDLLMessage = Left$(Puffer, RetVal)
End Function
