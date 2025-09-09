Attribute VB_Name = "movement"
' keyboard system
' thanks claude

#If Win64 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long
#End If

Private Const VK_SPACE As Long = 32
Private Const VK_W As Long = 87
Private Const VK_A As Long = 65
Private Const VK_S As Long = 83
Private Const VK_D As Long = 68
Private Const KEYEVENTF_KEYUP As Long = &H2

Private gameRunning As Boolean
#If Win64 Then
    Private timerID As LongPtr
#Else
    Private timerID As Long
#End If

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

Public Sub CancelWhiteout()
    keybd_event VK_W, 0, 0, 0
    keybd_event VK_W, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub CheckMovement()
    With ActivePresentation.Slides(1).Shapes("Player")
        If IsKeyPressed(VK_W) Then
            .IncrementTop (-10)
            CancelWhiteout
        End If
        If IsKeyPressed(VK_A) Then .IncrementLeft (-10)
        If IsKeyPressed(VK_S) Then .IncrementTop (10)
        If IsKeyPressed(VK_D) Then .IncrementLeft (10)
    End With
End Sub

Public Sub StartGame()
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
