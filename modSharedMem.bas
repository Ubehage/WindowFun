Attribute VB_Name = "modSharedMem"
Option Explicit

Global Const MSG_EXIT As Long = &H1
Global Const MSG_COLORFADE As Long = &H2
Global Const MSG_RUNNEXT As Long = &H3
Global Const MSG_SHRINKEXIT As Long = &H4
Global Const MSG_FADEEXIT As Long = &H5
Global Const MSG_SETBEEP As Long = &H6
Global Const MSG_ALLEFFECTS As Long = &H7

Enum AppMessages
  amExit = MSG_EXIT
  amColorFade = MSG_COLORFADE
  amRunNext = MSG_RUNNEXT
  amShrinkExit = MSG_SHRINKEXIT
  amFadeExit = MSG_FADEEXIT
  amSetBeep = MSG_SETBEEP
  amAllEffects = MSG_ALLEFFECTS
End Enum

Private Const PAGE_READWRITE As Long = &H4&
Private Const FILE_MAP_ALL_ACCESS As Long = &HF001F

Global Const SHAREDMEM_NAME As String = "Local\UbeWinFun23"
Private Const SHAREDMEM_DATASIZE As Long = 2
Private Const SHAREDMEM_SIZE As Long = 1024 * SHAREDMEM_DATASIZE

Public Type SHAREDMEM_DATA
  Data1 As Byte
  Data2 As Byte
End Type
Public Type SHARED_MEMORY_LAYOUT
  Level(0 To 1023) As SHAREDMEM_DATA
End Type

Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappingAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long

Private Declare Sub CopyMemoryByVal Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByVal lpSource As Any, ByVal cbCopy As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Dim SharedMemHandle As Long
Dim SharedMemBase As Long

Global SharedMemOffset As Long
Global SharedMemory As SHARED_MEMORY_LAYOUT

Public Function OpenSharedMemory() As Boolean
  Dim e As Boolean
  SharedMemHandle = CreateFileMapping(INVALID_HANDLE_VALUE, 0, PAGE_READWRITE, 0, SHAREDMEM_SIZE, SHAREDMEM_NAME)
  If SharedMemHandle = 0 Then Exit Function
  If Err.LastDllError() = ERROR_ALREADY_EXISTS Then e = True
  SharedMemBase = MapViewOfFile(SharedMemHandle, FILE_MAP_ALL_ACCESS, 0, 0, 0)
  If SharedMemBase = 0 Then
    Call CloseSharedMemory
    Exit Function
  End If
  If e = True Then
    Call ReadFromSharedMemory(True)
  Else
    ClearSharedMemory
  End If
  OpenSharedMemory = True
End Function

Public Function CloseSharedMemory() As Boolean
  Dim r As Boolean
  If SharedMemBase <> 0 Then
    Call UnmapViewOfFile(SharedMemBase)
    SharedMemBase = 0
    r = True
  End If
  If SharedMemHandle <> 0 Then
    Call CloseHandle(SharedMemHandle)
    SharedMemHandle = 0
    r = True
  End If
  CloseSharedMemory = r
End Function

Public Function ReadFromSharedMemory(Optional ReadAllMemory As Boolean = False, Optional LevelIndex As Long = 0) As Boolean
  If SharedMemBase = 0 Then Exit Function
  If ReadAllMemory = True Then
    CopyMemory SharedMemory, ByVal SharedMemBase, LenB(SharedMemory)
  Else
    Dim mAddr As Long, mOffset As Long
    mOffset = IIf(LevelIndex <= 0, SharedMemOffset, LevelIndex)
    If (mOffset < LBound(SharedMemory.Level) Or mOffset > UBound(SharedMemory.Level)) Then Exit Function
    mAddr = (SharedMemBase + (mOffset * SHAREDMEM_DATASIZE))
    CopyMemory SharedMemory.Level(mOffset), ByVal mAddr, LenB(SharedMemory.Level(mOffset))
  End If
  ReadFromSharedMemory = True
End Function

Public Function WriteToSharedMemory(Optional WriteAllMemory As Boolean = False, Optional LevelIndex As Long = 0) As Boolean
  If SharedMemBase = 0 Then Exit Function
  If WriteAllMemory = True Then
    CopyMemoryByVal SharedMemBase, SharedMemory, LenB(SharedMemory)
  Else
    Dim mAddr As Long, mOffset As Long
    mOffset = IIf(LevelIndex <= 0, SharedMemOffset, LevelIndex)
    If (mOffset < LBound(SharedMemory.Level) Or mOffset > UBound(SharedMemory.Level)) Then Exit Function
    mAddr = (SharedMemBase + (mOffset * SHAREDMEM_DATASIZE))
    CopyMemoryByVal mAddr, SharedMemory.Level(mOffset), LenB(SharedMemory.Level(mOffset))
  End If
  WriteToSharedMemory = True
End Function

Public Sub ClearSharedMemory()
  ZeroMemory SharedMemory, LenB(SharedMemory)
  Call WriteToSharedMemory(True)
End Sub
