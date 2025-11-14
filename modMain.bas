Attribute VB_Name = "modMain"
Option Explicit

'Windows transparency constants
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2

Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SETWINDOWPOS = SWP_NOSIZE Or SWP_NOMOVE

Private Const CMD_RUN = "bgnjdds"
Private Const CMD_LEVEL = "ubglzfz"
Private Const CMD_RANDOM = "ygrfjdrf"
Private Const CMD_SEP1 = ":"
Private Const CMD_SEP2 = ";"

Global Const MSG_SEP = ":"

Private Const REG_APPNAME = "WindowFun23"
Private Const REG_SETTINGS = "Settings"
Private Const REG_MESSAGE = "NextInst"

Global Const TIMER_INTERVAL As Long = 500
Global Const TIMER_INTERVAL_SHORT As Long = 50

Global Const INVALID_HANDLE_VALUE As Long = -1&
Global Const ERROR_ALREADY_EXISTS As Long = 183&

Type RGBColor
  Red As Integer
  Green As Integer
  Blue As Integer
End Type

Dim DoRun As Boolean
Dim RandomSeed As Long
Global AppLevel As Long
Global ExitLevel As Integer

Global StartLooping As Boolean
Global IsLooping As Boolean
Global DoColorOut As Boolean
Global DoShrink As Boolean
Global DoFadeOut As Boolean
Global FadeValue As Byte
Global DoAllEffects As Boolean

Global HideForm As Boolean
Global ShowForm As Boolean
Global NewRun As Boolean
Global UnloadNow As Boolean
Global NewWindow As Boolean

Global UnloadedByCode As Boolean
Global IsInIDE As Boolean

Global BeepWhenDone As Boolean
Global SetBeepState As Boolean

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Sub Main()
  IsInIDE = IDECheck()
  
  '#Debug
  'SplitCommandLine GetCommandLine(7)
  
  SplitCommandLine Command
  If RandomSeed = 0 Then RandomSeed = Timer
  Randomize RandomSeed
  If DoRun Then
    If Not OpenSharedMemory() = True Then
      MsgBox "Error opening shared memory!" & vbCrLf & "This program cannot continue."
      Exit Sub
    End If
    Load frmMain
    frmMain.SetForm
    frmMain.SetTimer
  Else
    ShowStartMessage
  End If
End Sub

Public Sub SetAppMessage(NewMessage As AppMessages)
  If NewMessage = amRunNext Then
    SharedMemory.Level(AppLevel).Data1 = AppMessages.amRunNext
    Call WriteToSharedMemory(False, AppLevel)
  Else
    Dim i As Long
    For i = LBound(SharedMemory.Level) To UBound(SharedMemory.Level)
      SharedMemory.Level(i).Data1 = NewMessage
    Next
    Call WriteToSharedMemory(True)
  End If
End Sub

Public Sub CheckAppMessage()
  If Not ReadFromSharedMemory(False, AppLevel) = True Then Exit Sub
  Select Case SharedMemory.Level(AppLevel).Data1
    Case AppMessages.amExit
      HideForm = True
      UnloadNow = True
      StartLooping = True
    Case AppMessages.amColorFade
      StartLooping = True
      DoColorOut = True
    Case AppMessages.amRunNext
      HideForm = True
      NewRun = True
      StartLooping = True
    Case AppMessages.amShrinkExit
      StartLooping = True
      DoShrink = True
    Case AppMessages.amFadeExit
      StartLooping = True
      DoFadeOut = True
      PrepareWindowTransparency
    Case AppMessages.amSetBeep
      SetBeepState = True
      BeepWhenDone = IIf(SharedMemory.Level(AppLevel).Data2 = 0, False, True)
    Case AppMessages.amAllEffects
      StartLooping = True
      DoAllEffects = True
      PrepareWindowTransparency
  End Select
End Sub

Public Sub ChangeBeepStateForAll()
  Dim i As Long
  For i = 1 To UBound(SharedMemory.Level)
    With SharedMemory.Level(i)
      .Data1 = AppMessages.amSetBeep
      .Data2 = IIf(BeepWhenDone = True, 1, 0)
    End With
  Next
  Call WriteToSharedMemory(True)
End Sub

Public Sub UnloadForm()
  UnloadedByCode = True
  Unload frmMain
  UnloadedByCode = False
  Set frmMain = Nothing
  If BeepWhenDone = True Then Beep
End Sub

Private Function GetAppMessage() As String
  Dim r As String
  On Error GoTo MsgError
  r = GetSetting(REG_APPNAME, REG_SETTINGS, REG_MESSAGE)
  On Error GoTo 0
  GetAppMessage = r
  Exit Function
MsgError:
  Resume Next
End Function

Private Sub SplitCommandLine(CommandLine As String)
  Dim i As Integer, cL() As String, cC() As String
  cL = Split(CommandLine, CMD_SEP2)
  For i = LBound(cL) To UBound(cL)
    cC = Split(cL(i), CMD_SEP1)
    Select Case cC(LBound(cC))
      Case CMD_RUN
        DoRun = True
      Case CMD_LEVEL
        AppLevel = CLng(cC(UBound(cC)))
      Case CMD_RANDOM
        RandomSeed = CLng(cC(UBound(cC)))
    End Select
  Next
End Sub

Private Function GetNextCommandLine(Optional Level As Long = -1) As String
  Dim aL As Long
  If Level = -1 Then aL = (AppLevel + 1) Else aL = Level
  GetNextCommandLine = """" & GetAppFile & """ " & GetCommandLine(aL)
End Function

Private Function GetCommandLine(NextLevel As Long) As String
  GetCommandLine = CMD_RUN & CMD_SEP2 & _
      CMD_LEVEL & CMD_SEP1 & Trim$(Str$(NextLevel)) & CMD_SEP2 & _
      CMD_RANDOM & CMD_SEP1 & Trim$(Str$(GetRandomNumber(1, 5000)))
End Function

Private Function GetAppFile() As String
  GetAppFile = FixPath(App.Path) & "\" & App.EXEName & ".exe"
End Function

Private Function FixPath(Path As String) As String
  Dim p As String
  p = Path
  While Right$(p, 1) = "\"
    p = Left$(p, (Len(p) - 1))
  Wend
  FixPath = Path
End Function

Private Sub RunNext()
  frmMain.Hide
  RunNew
  AppLevel = (AppLevel + 1)
  frmMain.SetRandomFormAndLabel
End Sub

Private Sub ShowStartMessage()
  Dim m As String
  m = "Instructions for WindowFun23:" & vbCrLf & vbCrLf
  m = m & "When you close the window, 2 new windows will pop up." & vbCrLf & _
          "When you close one window, all windows in that level will close and new ones will pop up. Lots of windows!" & vbCrLf & _
          "To close all windows and end the program, click inside one of the windows." & vbCrLf & vbCrLf & _
          "Function Keys:" & vbCrLf & _
          "Mouseclick - Close all windows and exit the program." & vbCrLf & _
          "ESC - Fades all windows to black and then closes." & vbCrLf & _
          "F - Shrinks all windows and then closes." & vbCrLf & vbCrLf & _
          "If you want to start this program, click 'Yes'. Otherwise, click 'No'."
  Select Case MsgBox(m, vbYesNo Or vbDefaultButton2, "WindowFun23 Instructions")
    Case vbYes
      Call OpenSharedMemory
      Call ClearSharedMemory
      
      
      'This is just for fun. Uncomment for a ride.
      'Dim i As Long, j As Long
      'j = 8
      'For i = 0 To j
      '  SharedMemory.Level(i).Data1 = AppMessages.amRunNext
      '  Call WriteToSharedMemory(False, i)
      'Next
      'SharedMemory.Level(j + 1).Data1 = AppMessages.amAllEffects
      'Call WriteToSharedMemory(False, (j + 1))
      
      
      RunNew 0
      CloseSharedMemory
  End Select
End Sub

Public Function GetRandomRGBColor() As RGBColor
  With GetRandomRGBColor
    .Red = GetRandomColorValue
    .Green = GetRandomColorValue
    .Blue = GetRandomColorValue
  End With
End Function

Public Function GetRandomColor() As Long
  GetRandomColor = RGBToColor(GetRandomRGBColor)
End Function

Public Function RGBToColor(ThisRGB As RGBColor) As Long
  With ThisRGB
    RGBToColor = RGB(.Red, .Green, .Blue)
  End With
End Function

Public Function InvertRGBColor(SourceRGB As RGBColor) As RGBColor
  With InvertRGBColor
    .Red = InvertColorValue(SourceRGB.Red)
    .Green = InvertColorValue(SourceRGB.Green)
    .Blue = InvertColorValue(SourceRGB.Blue)
  End With
End Function

Public Function CombineRGBColors(RGBColor1 As RGBColor, RGBColor2 As RGBColor) As RGBColor
  With CombineRGBColors
    .Red = CombineRGBValues(RGBColor1.Red, RGBColor2.Red)
    .Green = CombineRGBValues(RGBColor1.Green, RGBColor2.Green)
    .Blue = CombineRGBValues(RGBColor1.Blue, RGBColor2.Blue)
  End With
End Function

Private Function GetRandomColorValue() As Integer
  GetRandomColorValue = CInt(GetRandomNumber(0, 255))
End Function

Private Function InvertColorValue(SourceValue As Integer) As Integer
  Dim r As Integer
  r = SourceValue
  If r < 128 Then
    r = (128 + (128 - r))
  ElseIf r > 128 Then
    r = (128 - (r - 128))
  End If
  InvertColorValue = r
End Function

Private Function CombineRGBValues(RGBValue1 As Integer, RGBValue2 As Integer) As Integer
  If RGBValue1 > RGBValue2 Then
    CombineRGBValues = (RGBValue2 + ((RGBValue1 - RGBValue2) / 2))
  ElseIf RGBValue2 > RGBValue1 Then
    CombineRGBValues = (RGBValue1 + ((RGBValue2 - RGBValue1) / 2))
  Else
    CombineRGBValues = RGBValue1
  End If
End Function

Public Sub RunNew(Optional Level As Long = -1)
  On Error GoTo RunError
  Call Shell(GetNextCommandLine(Level), vbNormalFocus)
  On Error GoTo 0
  Exit Sub
RunError:
  Resume Next
End Sub

Public Function GetRandomNumber(Min As Long, Max As Long) As Long
  Dim r As Long
  r = ((Rnd * Max) + Min)
  If r < Min Then r = Min Else If r > Max Then r = Max
  GetRandomNumber = r
End Function

Public Sub PrepareWindowTransparency()
  Call SetWindowLong(frmMain.hWnd, GWL_EXSTYLE, GetWindowLong(frmMain.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
  FadeValue = 255
End Sub

Public Sub SetWindowTransparency(TransparencyLevel As Byte)
  Call SetLayeredWindowAttributes(frmMain.hWnd, 0, TransparencyLevel, LWA_ALPHA)
End Sub

Public Sub WindowOnTop(hWnd As Long, OnTop As Boolean)
  Dim wFlags As Long
  If OnTop Then
    wFlags = HWND_TOPMOST
  Else
    wFlags = HWND_NOTOPMOST
  End If
  SetWindowPos hWnd, wFlags, 0&, 0&, 0&, 0&, SWP_SETWINDOWPOS
End Sub

Private Function IDECheck() As Boolean
  Dim iVar As Boolean
  On Error GoTo IDEError
  Debug.Assert TestIDE(iVar)
  On Error GoTo 0
  IDECheck = iVar
  Exit Function
IDEError:
  Resume Next
End Function

Private Function TestIDE(TestVar As Boolean) As Boolean
  TestVar = True
  TestIDE = True
End Function
