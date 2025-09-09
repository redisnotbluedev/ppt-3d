VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Render3D 
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "Render3D.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Render3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Borderless window
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

' reg
Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
    ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As LongPtr) As Long

Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
    ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" ( _
    ByVal hKey As LongPtr) As Long

Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_DWORD = 4

Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_FRAMECHANGED = &H20

Public Function ReadFileToString(fileName As String) As String
    Dim filePath As String
    Dim fileNumber As Integer
    Dim fileContent As String
    
    ' Build full path (assumes file is in same folder as workbook)
    filePath = ActivePresentation.Path & "\" & fileName
    
    ' Check if file exists
    If Dir(filePath) = "" Then
        ReadFileToString = ""
        Debug.Print "File not found: " & filePath
        Exit Function
    End If
    
    ' Read the entire file in one go
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    fileContent = Input$(LOF(fileNumber), fileNumber)
    Close #fileNumber
    
    ReadFileToString = fileContent
End Function

' Method 1: More aggressive registry approach with multiple keys
Public Sub ForceIE11Nuclear()
    Dim hKey As LongPtr
    Dim lResult As Long
    Dim appNames() As String
    Dim i As Integer
    
    ' Try multiple app name variations
    appNames = Split("POWERPNT.EXE,powerpnt.exe,PowerPoint,EXCEL.EXE,VBE7.EXE,VBE6.EXE", ",")
    
    For i = 0 To UBound(appNames)
        ' FEATURE_BROWSER_EMULATION
        SetRegistryValue "SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION", appNames(i), 11001
        
        ' FEATURE_DISABLE_LEGACY_COMPRESSION
        SetRegistryValue "SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_DISABLE_LEGACY_COMPRESSION", appNames(i), 1
        
        ' FEATURE_LOCALMACHINE_LOCKDOWN
        SetRegistryValue "SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_LOCALMACHINE_LOCKDOWN", appNames(i), 0
        
        ' FEATURE_BLOCK_LMZ_SCRIPT
        SetRegistryValue "SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BLOCK_LMZ_SCRIPT", appNames(i), 0
    Next i
    
    Debug.Print "Applied nuclear registry settings"
End Sub

Private Sub SetRegistryValue(keyPath As String, valueName As String, valueData As Long)
    Dim hKey As LongPtr
    Dim lResult As Long
    
    lResult = RegOpenKeyEx(HKEY_CURRENT_USER, keyPath, 0, KEY_ALL_ACCESS, hKey)
    If lResult = 0 Then
        RegSetValueEx hKey, valueName, 0, REG_DWORD, valueData, 4
        RegCloseKey hKey
        Debug.Print "Set " & keyPath & "\" & valueName & " = " & valueData
    End If
End Sub

Private Sub UserForm_Activate()
    ForceIE11Nuclear
    ResizeUserForm
    
    Dim htmlContent As String
    htmlContent = ReadFileToString("renderer.html")
    
    WB_Game.Navigate "about:blank"
    
    ' Wait for navigation
    Dim t As Single
    t = Timer
    Do While Timer < t + 0.5
        DoEvents
    Loop
    
    On Error Resume Next
    WB_Game.Document.Open
    WB_Game.Document.Write htmlContent
    WB_Game.Document.Close
    On Error GoTo 0
End Sub

Private Sub ResizeUserForm()
    Dim slideWidth As Single, slideHeight As Single
    Dim pptWidth As Single, pptHeight As Single
    Dim pptLeft As Single, pptTop As Single
    Dim actualWidth As Single, actualHeight As Single
    Dim leftOffset As Single, topOffset As Single
    Dim aspectRatio As Single
    Dim margin As Single
    
    ' Set margin size (in points)
    margin = 20  ' Adjust this value as needed
    
    ' Get slide dimensions and aspect ratio
    slideWidth = ActivePresentation.PageSetup.slideWidth
    slideHeight = ActivePresentation.PageSetup.slideHeight
    aspectRatio = slideWidth / slideHeight
    
    ' Get PowerPoint window dimensions
    On Error Resume Next
    If ActivePresentation.SlideShowSettings.ShowType = ppShowTypeSpeaker Or _
       ActivePresentation.SlideShowSettings.ShowType = ppShowTypeKiosk Then
        ' In slide show mode
        Dim ssWindow As SlideShowWindow
        Set ssWindow = ActivePresentation.SlideShowWindow
        If Not ssWindow Is Nothing Then
            pptWidth = ssWindow.Width
            pptHeight = ssWindow.Height
            pptLeft = ssWindow.Left
            pptTop = ssWindow.Top
        End If
    Else
        ' Normal presentation mode
        Dim presWindow As DocumentWindow
        Set presWindow = ActiveWindow
        If Not presWindow Is Nothing Then
            pptWidth = presWindow.Width
            pptHeight = presWindow.Height
            pptLeft = presWindow.Left
            pptTop = presWindow.Top
        End If
    End If
    On Error GoTo 0
    
    ' Fallback to screen dimensions if PowerPoint window not found
    If pptWidth = 0 Or pptHeight = 0 Then
        pptWidth = GetSystemMetrics(0) * 0.75
        pptHeight = GetSystemMetrics(1) * 0.75
        pptLeft = 0
        pptTop = 0
    End If
    
    ' Calculate actual content area within PowerPoint window (accounting for letterbox/pillarbox)
    If pptWidth / pptHeight > aspectRatio Then
        ' Letterboxed (black bars on left/right)
        actualHeight = pptHeight
        actualWidth = pptHeight * aspectRatio
        leftOffset = pptLeft + (pptWidth - actualWidth) / 2
        topOffset = pptTop
    Else
        ' Pillarboxed (black bars on top/bottom)
        actualWidth = pptWidth
        actualHeight = pptWidth / aspectRatio
        leftOffset = pptLeft
        topOffset = pptTop + (pptHeight - actualHeight) / 2
    End If
    
    ' Apply margin to content area
    actualWidth = actualWidth - (margin * 2)
    actualHeight = actualHeight - (margin * 2)
    leftOffset = leftOffset + margin
    topOffset = topOffset + margin
    
    ' Make UserForm truly borderless
    Me.BorderStyle = fmBorderStyleNone
    Me.Caption = ""
    
    ' Size to actual content area with margin
    Me.Width = actualWidth
    Me.Height = actualHeight
    Me.Left = leftOffset
    Me.Top = topOffset
    
    ' Remove title bar using Windows API
    RemoveTitleBar
    
    ' Resize browser control to fill entire form
    WB_Game.Left = 0
    WB_Game.Top = 0
    WB_Game.Width = actualWidth
    WB_Game.Height = actualHeight
End Sub

Private Sub UserForm_Initialize()
    ' Set borderless on initialize
    Me.BorderStyle = fmBorderStyleNone
    Me.Caption = ""
End Sub

Private Sub RemoveTitleBar()
    Dim hWnd As LongPtr
    Dim lStyle As Long
    
    ' Get the window handle for this UserForm
    hWnd = FindWindow("ThunderDFrame", Me.Caption)
    If hWnd = 0 Then hWnd = FindWindow(vbNullString, Me.Caption)
    
    If hWnd <> 0 Then
        ' Get current window style and remove caption
        lStyle = GetWindowLong(hWnd, GWL_STYLE)
        lStyle = lStyle And Not WS_CAPTION
        SetWindowLong hWnd, GWL_STYLE, lStyle
        
        ' Apply the changes
        SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED
    End If
End Sub

Private Sub WB_Game_StatusTextChange(ByVal Text As String)
End Sub

