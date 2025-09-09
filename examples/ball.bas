Attribute VB_Name = "Module1"
' Windows API declarations
#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
#Else
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
#End If

' Game variables
Private gameRunning As Boolean
#If VBA7 Then
    Private timerID As LongPtr
#Else
    Private timerID As Long
#End If
Private velocity As Single
Private Const GRAVITY As Single = 0.3
Private Const BOUNCE_DAMPING As Single = 0.7

Public Sub StartFalling()
    gameRunning = True
    velocity = 0 ' Start with no velocity
    
    ' Create the falling shape if it doesn't exist
    On Error Resume Next
    ActivePresentation.Slides(1).Shapes("FallingShape").Delete
    On Error GoTo 0
    
    With ActivePresentation.Slides(1).Shapes.AddShape(msoShapeOval, 200, 50, 40, 40)
        .Name = "FallingShape"
        .Fill.ForeColor.RGB = RGB(255, 100, 100)
        .Line.Weight = 2
        .Line.ForeColor.RGB = RGB(200, 50, 50)
    End With
    
    ' Start the timer (30ms for smooth physics)
    timerID = SetTimer(GetActiveWindow(), 0, 50, AddressOf FallCallback)
    
    If timerID = 0 Then
        MsgBox "Timer failed to start!"
        gameRunning = False
    End If
End Sub

Public Sub StopFalling()
    gameRunning = False
    If timerID <> 0 Then
        KillTimer GetActiveWindow(), timerID
        timerID = 0
    End If
End Sub

' Timer callback - physics simulation
#If VBA7 Then
    Private Sub FallCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
    Private Sub FallCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    On Error GoTo StopOnError
    
    If Not gameRunning Then Exit Sub
    
    Dim fallingShape As Shape
    Set fallingShape = ActivePresentation.Slides(1).Shapes("FallingShape")
    
    ' Apply gravity
    velocity = velocity + GRAVITY
    
    ' Calculate new position
    Dim newTop As Single
    newTop = fallingShape.Top + velocity
    
    ' Check for ground collision
    Dim slideHeight As Single
    slideHeight = ActivePresentation.PageSetup.slideHeight
    Dim shapeBottom As Single
    shapeBottom = newTop + fallingShape.Height
    
    If shapeBottom >= slideHeight Then
        ' Hit the ground - stop and bounce slightly
        newTop = slideHeight - fallingShape.Height
        velocity = -velocity * BOUNCE_DAMPING
        
        ' Stop bouncing if velocity is too small
        If Abs(velocity) < 1 Then
            velocity = 0
        End If
    End If
    
    ' Check collision with other shapes
    If CheckShapeCollision(fallingShape, newTop) Then
        ' Hit another shape - stop falling
        velocity = -velocity * BOUNCE_DAMPING * 0.5 ' Smaller bounce off shapes
        If Abs(velocity) < 0.5 Then
            velocity = 0
        End If
    Else
        ' No collision - move to new position
        fallingShape.Top = newTop
    End If
    
    Exit Sub
    
StopOnError:
    StopFalling
End Sub

' Check if the falling shape would collide with other shapes
Private Function CheckShapeCollision(fallingShape As Shape, newTop As Single) As Boolean
    CheckShapeCollision = False
    
    Dim slide As slide
    Set slide = ActivePresentation.Slides(1)
    
    Dim otherShape As Shape
    Dim fallingLeft As Single, fallingRight As Single
    Dim fallingBottom As Single, otherLeft As Single
    Dim otherRight As Single, otherTop As Single, otherBottom As Single
    
    fallingLeft = fallingShape.Left
    fallingRight = fallingShape.Left + fallingShape.Width
    fallingBottom = newTop + fallingShape.Height
    
    For Each otherShape In slide.Shapes
        ' Skip the falling shape itself
        If otherShape.Name <> "FallingShape" Then
            otherLeft = otherShape.Left
            otherRight = otherShape.Left + otherShape.Width
            otherTop = otherShape.Top
            otherBottom = otherShape.Top + otherShape.Height
            
            ' Check if rectangles overlap
            If fallingRight > otherLeft And fallingLeft < otherRight And _
               fallingBottom > otherTop And newTop < otherBottom Then
                ' Collision detected - position shape just above the other shape
                fallingShape.Top = otherTop - fallingShape.Height
                CheckShapeCollision = True
                Exit Function
            End If
        End If
    Next otherShape
End Function

' Helper function to add obstacles
Public Sub AddObstacle()
    Static obstacleCount As Integer
    obstacleCount = obstacleCount + 1
    
    With ActivePresentation.Slides(1).Shapes.AddShape(msoShapeRectangle, 100 + (obstacleCount * 80), 300, 60, 20)
        .Name = "Obstacle" & obstacleCount
        .Fill.ForeColor.RGB = RGB(100, 100, 255)
    End With
End Sub

