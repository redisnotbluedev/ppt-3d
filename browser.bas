Attribute VB_Name = "browser"
' keyboard system
' thanks claude

#If Win64 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
#End If
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const VK_SPACE As Long = 32
Private Const VK_W As Long = 87
Private Const VK_A As Long = 65
Private Const VK_S As Long = 83
Private Const VK_D As Long = 68
Private Const VK_Q As Long = &HA0
Private Const VK_E As Long = &H20
Private Const SENS As Double = 0.005

Private gameRunning As Boolean
#If Win64 Then
    Private timerID As LongPtr
#Else
    Private timerID As Long
#End If

Dim lastX As Long, lastY As Long

Public Sub GetMousePosition(ByRef x As Long, ByRef y As Long)
    Dim p As POINTAPI
    GetCursorPos p
    x = p.x
    y = p.y
End Sub

Public Function IsKeyPressed(keyCode As Long) As Boolean
    IsKeyPressed = (GetAsyncKeyState(keyCode) And &H8000) <> 0
End Function

Public Function IsKeyJustPressed(keyCode As Long) As Boolean
    Static keyStates(255) As Boolean
    Dim currentState As Boolean
    currentState = IsKeyPressed(keyCode)
    
    If keyCode < 0 Or keyCode > 255 Then
        IsKeyJustPressed = False
        Exit Function
    End If
    
    IsKeyJustPressed = currentState And Not keyStates(keyCode)
    keyStates(keyCode) = currentState
End Function

Private Sub CancelEffects()
    Dim ssw As SlideShowWindow
    On Error Resume Next
    Set ssw = ActivePresentation.SlideShowWindow
    If Not ssw Is Nothing Then
        ssw.View.State = ppSlideShowRunning
    End If
End Sub

Private Sub UpdateMouseLook()
    ShowCursor 0
    Static centerX As Long, centerY As Long
    Static initialized As Boolean
    
    If Not initialized Then
        Dim windowRect As RECT
        Dim hwnd As LongPtr
        hwnd = GetActiveWindow()
        
        GetWindowRect hwnd, windowRect
        centerX = (windowRect.Left + windowRect.Right) / 2
        centerY = (windowRect.Top + windowRect.Bottom) / 2
        
        SetCursorPos centerX, centerY
        initialized = True
        Exit Sub
    End If
    
    Dim mouseX As Long, mouseY As Long
    GetMousePosition mouseX, mouseY
    
    Dim dx As Long, dy As Long
    dx = mouseX - centerX
    dy = mouseY - centerY
    
    If dx <> 0 Or dy <> 0 Then
        Render3D.WB_Game.Document.parentWindow.rotateCamera -dy * SENS, dx * SENS, 0
        SetCursorPos centerX, centerY
    End If
End Sub

Private Sub CheckMovement()
    With Render3D.WB_Game.Document.parentWindow
        If IsKeyPressed(VK_W) Then .moveCamera 0, 0, 10
        If IsKeyPressed(VK_A) Then .moveCamera -10, 0, 0
        If IsKeyPressed(VK_S) Then .moveCamera 0, 0, -10
        If IsKeyPressed(VK_D) Then .moveCamera 10, 0, 0
        If IsKeyPressed(VK_Q) Then .moveCamera 0, -10, 0
        If IsKeyPressed(VK_E) Then .moveCamera 0, 10, 0
        
        UpdateMouseLook
    End With
End Sub

Private Sub StartGame()
    gameRunning = True
    timerID = SetTimer(GetActiveWindow(), 0, 30, AddressOf GameLoop)
    
    If timerID = 0 Then
        MsgBox "Timer failed to start!"
        gameRunning = False
    End If
End Sub

Public Sub EndGame()
    gameRunning = False
    If timerID <> 0 Then
        KillTimer GetActiveWindow(), timerID
        timerID = 0
    End If
End Sub

#If Win64 Then
    Private Sub GameLoop(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
    Private Sub GameLoop(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    On Error Resume Next
    If Not gameRunning Then Exit Sub
    
    Debug.Print "Loop running"
    CheckMovement
    If Err.Number <> 0 Then
        Debug.Print "Error: " & Err.Description
        Err.Clear
    End If
End Sub

Sub LaunchHelloWorld()
    GetMousePosition lastX, lastY
    Render3D.Show
    StartGame
End Sub
