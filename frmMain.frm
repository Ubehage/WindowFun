VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label l2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   285
      Width           =   645
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LABEL_CAPTION = "Click - Close and end.%b%" & _
                              "Esc - Fade to black and close.%b%" & _
                              "F - Shrink and close.%b%" & _
                              "C - Fade out and close.%b%" & _
                              "A - Do all effects.%b%" & _
                              "B - Beep when closing: %beep%"

Private Const COLOR_FADE_VALUE = 5
Private Const LABEL_SIZE_CHANGE_SMALL = 2
Private Const LABEL_SIZE_CHANGE_LARGE = 100
Private Const WINDOW_FADE_VALUE = 10

Dim WithEvents fTimer As WindowTimer
Attribute fTimer.VB_VarHelpID = -1

Dim FormColor As RGBColor
Dim LabelColor As RGBColor
Dim FormText As String

Dim FormLeft As Long, FormTop As Long
Dim FormWidth As Long, FormHeight As Long

Dim ScaleDiffX As Long
Dim ScaleDiffY As Long

Friend Sub SetForm()
  ScaleDiffX = (Me.Width - Me.ScaleWidth)
  ScaleDiffY = (Me.Height - Me.ScaleHeight)
  WindowOnTop Me.hWnd, True
  Me.AutoRedraw = True
  Me.Font.Size = modMain.GetRandomNumber(59, 521)
  SetRandomFormAndLabel
End Sub

Friend Sub SetTimer()
  Set fTimer = New WindowTimer
  With fTimer
    .Interval = TIMER_INTERVAL
    .Enabled = True
  End With
End Sub

Friend Sub KillTimer()
  If Not fTimer Is Nothing Then
    fTimer.Enabled = False
    Set fTimer = Nothing
  End If
End Sub

Friend Sub SetNewTimerInterval(NewInterval As Long)
  If Not fTimer Is Nothing Then fTimer.Interval = NewInterval
End Sub

Friend Sub SetRandomFormAndLabel(Optional DoNotShow As Boolean = False, Optional SetLabel As Boolean = True)
  If SetLabel Then
    l2.Move 15, 15
    SetLabelCaption
    l2.Visible = True
    SetFormText True, True, FormWidth, FormHeight, False
  End If
  FormLeft = GetRandomNumber(0, (Screen.Width - FormWidth))
  FormTop = GetRandomNumber(0, (Screen.Height - FormHeight))
  SetFormColor
  Me.Caption = CStr(AppLevel)
  MoveForm FormLeft, FormTop, FormWidth, FormHeight, SetLabel
  If (Me.Visible = False And DoNotShow = False) Then
    Me.Show
  End If
End Sub

Private Sub SetLabelCaption()
  l2.Caption = GetLabelCaption
End Sub

Private Function GetLabelCaption() As String
  GetLabelCaption = Replace(Replace(LABEL_CAPTION, "%b%", vbCrLf), "%beep%", IIf(BeepWhenDone = True, "ON", "OFF"))
End Function

Private Sub SetFormSize()
  SetFormText True, True, FormWidth, FormHeight, False
End Sub

Private Sub SetRandomFormSize()
  FormWidth = GetRandomNumber((Screen.Width * 0.05), (Screen.Width * 0.7))
  FormHeight = GetRandomNumber((Screen.Height * 0.05), (Screen.Height * 0.7))
End Sub

Private Sub MoveForm(LeftPos As Long, TopPos As Long, Width As Long, Height As Long, Optional DrawText As Boolean = False)
  Me.Move LeftPos, TopPos, (ScaleDiffX + Width), (ScaleDiffY + Height)
  If DrawText Then
    SetFormText False, False, 0, 0, True
  End If
End Sub

Private Sub SetFormText(SetCaption As Boolean, SetFont As Boolean, Width As Long, Height As Long, DoDraw As Boolean)
  If SetCaption = True Then FormText = GetFormText()
  If SetFont Then
    With Me.Font
      .Name = Screen.Fonts(GetRandomNumber(0, (Screen.FontCount - 1)))
      .Bold = True
    End With
  End If
  If DoDraw Then
    Me.CurrentX = 15
    Me.CurrentY = 15
    Me.Print FormText
  Else
    Width = (Me.TextWidth(FormText) + 30)
    If Me.TextHeight(FormText) < l2.Height Then
      Height = (l2.Height + 30)
    Else
      Height = (Me.TextHeight(FormText) + 30)
    End If
  End If
End Sub

Private Sub SetTextSize(Optional SetCaption As Boolean = False, Optional SetFont As Boolean = False, Optional ChangeInterval As Integer = LABEL_SIZE_CHANGE_LARGE)
  Dim i As Integer, tW As Long, tH As Long
  If SetCaption Then FormText = GetFormText
  If SetFont Then
    With Me.Font
      .Name = Screen.Fonts(GetRandomNumber(0, (Screen.FontCount - 1)))
      .Bold = True
    End With
  End If
  GoTo AfterLoop
  
  i = ChangeInterval
  On Error GoTo TextError
  GoSub SetSizeVars
  If (tW > FormWidth Or tH > FormHeight) Then
    Do
      Me.Font.Size = (Me.Font.Size - i)
      GoSub SetSizeVars
      If (tW <= FormWidth And tH <= FormHeight) Then
        Me.Font.Size = (Me.Font.Size + i)
        GoSub CheckI
      End If
    Loop
  ElseIf (tW < FormWidth Or tH < FormHeight) Then
    Do
      Me.Font.Size = (Me.Font.Size + i)
      GoSub SetSizeVars
      If (tW > FormWidth Or tH > FormHeight) Then
        Me.Font.Size = (Me.Font.Size - i)
        GoSub CheckI
      End If
    Loop
  End If
AfterLoop:
  Me.Cls
  Me.CurrentX = ((Me.ScaleWidth - tW) / 2)
  Me.CurrentY = ((Me.ScaleHeight - tH) / 2)
  Me.Print FormText
  Exit Sub
SetSizeVars:
  tW = Me.TextWidth(FormText)
  tH = Me.TextHeight(FormText)
  Return
CheckI:
  If i = 1 Then GoTo AfterLoop
  i = (i / 2)
  Return
TextError:
  Resume AfterLoop
End Sub

Private Function GetFormText() As String
  Dim i As Integer, c As Long
  i = 0
  c = 1
  Do
    If i >= AppLevel Then Exit Do
    i = (i + 1)
    c = (c * 2)
  Loop
  GetFormText = CStr(c)
End Function

Private Sub SetFormColor()
  With GetRandomRGBColor
    FormColor.Red = .Red
    FormColor.Green = .Green
    FormColor.Blue = .Blue
  End With
  LabelColor = InvertRGBColor(FormColor)
  UpdateColors
End Sub

Private Sub UpdateColors()
  Me.BackColor = RGBToColor(FormColor)
  Me.ForeColor = RGBToColor(LabelColor)
  l2.ForeColor = RGBToColor(LabelColor)
End Sub

Private Function DoColorFade(Optional FadeValue As Integer = COLOR_FADE_VALUE) As Boolean
  Dim r As Boolean
  If FadeThisColor(FormColor, FadeValue) Then r = True
  If FadeThisColor(LabelColor, FadeValue) Then r = True
  If r = True Then
    UpdateColors
    SetFormText False, False, 0, 0, True
  End If
  DoColorFade = r
End Function

Private Function FadeThisColor(ThisColor As RGBColor, Optional FadeValue As Integer = COLOR_FADE_VALUE) As Boolean
  Dim r As Boolean
  With ThisColor
    If FadeThisColorValue(.Red, FadeValue) Then r = True
    If FadeThisColorValue(.Green, FadeValue) Then r = True
    If FadeThisColorValue(.Blue, FadeValue) Then r = True
  End With
  FadeThisColor = r
End Function

Private Function FadeThisColorValue(ThisValue As Integer, Optional FadeValue As Integer = COLOR_FADE_VALUE) As Boolean
  If ThisValue > 0 Then
    ThisValue = (ThisValue - FadeValue)
    If ThisValue < 0 Then ThisValue = 0
    FadeThisColorValue = True
  End If
End Function

Private Function DoTheShrink() As Boolean
  Dim tX As Long, tY As Long
  If Me.Font.Size <= 15 Then DoTheShrink = True: Exit Function
  Me.Font.Size = (Me.Font.Size - 10)
  SetFormText False, False, tX, tY, False
  FormLeft = (FormLeft + ((FormWidth - tX) \ 2))
  FormTop = (FormTop + ((FormHeight - tY) \ 2))
  FormWidth = tX
  FormHeight = tY
  Me.Cls
  MoveForm FormLeft, FormTop, FormWidth, FormHeight, True
End Function

Private Function DoTheFadeOut() As Boolean
  If FadeValue = 0 Then
    DoTheFadeOut = True
  Else
    Dim v As Single
    v = (FadeValue - WINDOW_FADE_VALUE)
    If v < 0 Then FadeValue = 0 Else FadeValue = v
    SetWindowTransparency FadeValue
  End If
End Function

Private Function DoAllTheEffects() As Boolean
  Dim r As Boolean
  If DoColorFade() Then r = True
  If r = True Then If DoTheShrink() Then r = False
  If r = True Then If DoTheFadeOut() Then r = False
  DoAllTheEffects = r
End Function

Private Sub ToggleBeepWhenDone()
  BeepWhenDone = Not BeepWhenDone
  ChangeBeepStateForAll
End Sub

Private Sub Form_Click()
  If IsLooping Then
    UnloadNow = True
  Else
    SetAppMessage amExit
  End If
End Sub

Friend Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not IsLooping Then
    Select Case KeyCode
      Case vbKeyEscape
        SetAppMessage amColorFade
      Case vbKeyF
        SetAppMessage amShrinkExit
      Case vbKeyC
        SetAppMessage amFadeExit
      Case vbKeyB
        ToggleBeepWhenDone
      Case vbKeyA
        SetAppMessage amAllEffects
    End Select
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If UnloadedByCode Then
    KillTimer
    WindowOnTop Me.hWnd, False
    Call CloseSharedMemory
  Else
    If Not DoColorOut Then
      SetAppMessage amRunNext
      Cancel = 1
    End If
  End If
End Sub

Private Sub fTimer_Timer()
  fTimer.Enabled = False
  If SetBeepState = True Then
    SetLabelCaption
    SetBeepState = False
  ElseIf StartLooping Then
    IsLooping = StartLooping
    StartLooping = Not StartLooping
    SetNewTimerInterval TIMER_INTERVAL_SHORT
    l2.Visible = False
  ElseIf IsLooping Then
    If UnloadNow Then
      If Me.Visible = True Then
        If HideForm = True Then GoTo NextTimerStep
        HideForm = True
      Else
        UnloadForm
      End If
      GoTo ExitTimer
    End If
NextTimerStep:
    If (ShowForm = True Or HideForm = True) Then
      If ShowForm = True Then
        Me.Show
        ShowForm = False
      End If
      If HideForm = True Then
        Me.Hide
        HideForm = False
      End If
    Else
      If Me.Visible = False Then
        If NewWindow = True Then
          SetRandomFormAndLabel True
          NewWindow = False
          ShowForm = True
        End If
        If NewRun = True Then
          RunNew
          AppLevel = (AppLevel + 1)
          NewRun = False
          NewWindow = True
        End If
      Else
        If DoColorOut Then
          If Not DoColorFade Then UnloadNow = True
        ElseIf DoShrink Then
          If DoTheShrink Then UnloadNow = True
        ElseIf DoFadeOut Then
          If DoTheFadeOut Then UnloadNow = True
        ElseIf DoAllEffects Then
          If DoAllTheEffects() = False Then UnloadNow = True
        Else
          IsLooping = False
          SetNewTimerInterval TIMER_INTERVAL
        End If
      End If
    End If
  Else
    CheckAppMessage
  End If
ExitTimer:
  If Not fTimer Is Nothing Then fTimer.Enabled = True
End Sub

Private Sub l2_Click()
  Call Form_Click
End Sub
