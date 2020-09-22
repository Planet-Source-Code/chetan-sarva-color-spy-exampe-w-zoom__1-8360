VERSION 5.00
Begin VB.Form frmZoom 
   BorderStyle     =   0  'None
   Caption         =   "Screen Zoom"
   ClientHeight    =   1995
   ClientLeft      =   4110
   ClientTop       =   3720
   ClientWidth     =   1995
   Icon            =   "frmZoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   133
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrZoom 
      Interval        =   50
      Left            =   1710
      Top             =   360
   End
   Begin VB.PictureBox picZoom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   780
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   0
      Top             =   660
      Width           =   945
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' Code by: Rocky Clark (Kath-Rock Software)
'
' Zooming "crosshair" added by: Chetan Sarva



'User Defined Types
Private Type PointAPI   'API point structure.
    X   As Long
    Y   As Long
End Type

Private Type SizeRect   'Size structure (uses Width, Height instead of bounds)
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
End Type

Private Type RectAPI    'Rect structure (uses Right, Bottom bounds instead of Width, Height)
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'Windows API Blt (BitBlt, PatBlt, StretchBlt) ROP constants.
Private Const SRCCOPY           As Long = &HCC0020
Private Const PATCOPY           As Long = &HF00021

'SetWindowPos Flags.
Private Const SWP_NOMOVE        As Long = 2
Private Const SWP_NOSIZE        As Long = 1
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_FLAGS         As Long = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
Private Const HWND_TOPMOST      As Long = -1
Private Const HWND_NOTOPMOST    As Long = -2

'Module level variables.
Private mfScale As Single   'Scale of Zoom percentage (6 = 600%) (6 x Size = 600% increase)
Private mlOldX  As Long     'Holds Last X-coord of mouse
Private mlOldY  As Long     'Holds Last Y-coord of mouse

'Declare the Windows API functions that are to be used.
'Alphabetical order to ease lookup later.
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RectAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Function CreateCheckeredBrush(ByVal hDC As Long, ByVal lColor1 As Long, ByVal lColor2 As Long) As Long

Dim X           As Long
Dim Y           As Long
Dim lRet        As Long
Dim hBitmapDC   As Long
Dim hBitmap     As Long
Dim hOldBitmap  As Long
    
    'Convert System Colors if needed
    If lColor1 < 0 Then
        lColor1 = GetSysColor(lColor1 And &HFF&)
    End If
    If lColor2 < 0 Then
        lColor2 = GetSysColor(lColor2 And &HFF&)
    End If
    
    'Create a new DC and Bitmap to draw the Brush
    hBitmapDC = CreateCompatibleDC(hDC)
    hBitmap = CreateCompatibleBitmap(hDC, 8, 8)
    'Select the Bitmap into the DC for drawing
    hOldBitmap = SelectObject(hBitmapDC, hBitmap)
    
    'Draw the Brush's Bitmap (Checkerboard)
    For Y = 0 To 6 Step 2
        For X = 0 To 6 Step 2
            lRet = SetPixelV(hBitmapDC, X, Y, lColor1)
            lRet = SetPixelV(hBitmapDC, X + 1, Y, lColor2)
            lRet = SetPixelV(hBitmapDC, X, Y + 1, lColor2)
            lRet = SetPixelV(hBitmapDC, X + 1, Y + 1, lColor1)
        Next X
    Next Y
    
    'Get the bitmap back out of the DC
    hBitmap = SelectObject(hBitmapDC, hOldBitmap)
    
    'Create the Brush from the bitmap
    CreateCheckeredBrush = CreatePatternBrush(hBitmap)
    
    'Delete the DC and Bitmap to free memory
    lRet = DeleteDC(hBitmapDC)
    lRet = DeleteObject(hBitmap)

End Function

Private Sub DoZoom(ptMouse As PointAPI)

Dim lRet        As Long
Dim lTemp       As Long
Dim hWndDesk    As Long
Dim hDCDesk     As Long
Dim sizSrce     As SizeRect
Dim sizDest     As SizeRect

    'Get the Desktop DC
    hWndDesk = GetDesktopWindow()
    hDCDesk = GetDC(hWndDesk)
    
    'Setup the Destination size for StretchBlt.
    With sizDest
        .Left = 0
        .Top = 0
        .Width = picZoom.ScaleWidth
        .Height = picZoom.ScaleHeight
    End With
    
    'Setup the Source size for StretchBlt.
    With sizSrce
        .Left = ptMouse.X - Int((sizDest.Width / 2) / mfScale)
        .Top = ptMouse.Y - Int((sizDest.Height / 2) / mfScale)
        .Width = Int(sizDest.Width / mfScale)
        .Height = Int(sizDest.Height / mfScale)
        'Adjust Source and Destination sizes if they don't match.
        'sizSrce.Size * mfScale must= sizDest.Size for acurate scaling.
        'Destination must always be as large or larger than picZoom.
        'Adjust the Width, if needed.
        lTemp = Int(.Width * mfScale)  '(Source.Width * mfScale must= sizDest.Width)
        If lTemp > sizDest.Width Then
            sizDest.Width = lTemp
        ElseIf lTemp < sizDest.Width Then
            .Width = .Width + 1
            sizDest.Width = lTemp + mfScale
        End If
        'Adjust the Height, if needed.
        lTemp = Int(.Height * mfScale) '(sizSrce.Height * mfScale must= sizDest.Height)
        If lTemp > sizDest.Height Then
            sizDest.Height = lTemp
        ElseIf lTemp < sizDest.Height Then
            .Height = .Height + 1
            sizDest.Height = lTemp + mfScale
        End If
    End With
    
    'Clear the current contents.
    picZoom.Cls
    
    'Stretch the Desktop (source) into picZoom (dest)
    lRet = StretchBlt(picZoom.hDC, sizDest.Left, sizDest.Top, sizDest.Width, sizDest.Height, hDCDesk, sizSrce.Left, sizSrce.Top, sizSrce.Width, sizSrce.Height, SRCCOPY)
    
    'Release the Desktop DC
    lRet = ReleaseDC(hWndDesk, hDCDesk)
    
    'Redraw the grid
    Call DrawGrid

    
    picZoom.Refresh
    
End Sub

Private Sub DrawGrid()

Dim iWidth      As Integer
Dim iHeight     As Integer
Dim lRet        As Long
Dim hBrush      As Long
Dim hOldBrush   As Long
Dim fX          As Single
Dim fY          As Single

    If mfScale >= 3 Then
    
        'Create a Checkered Brush (Dark and Light Grey)...
        hBrush = CreateCheckeredBrush(picZoom.hDC, &H808080, &HC0C0C0)
        '...and Select it into the PictureBox
        hOldBrush = SelectObject(picZoom.hDC, hBrush)
        
        iWidth = picZoom.ScaleWidth
        iHeight = picZoom.ScaleHeight
        
        'Draw the gridlines using the checkered pattern brush.
        For fX = 0 To iWidth Step mfScale
            lRet = PatBlt(picZoom.hDC, Int(fX), 0, 1, iHeight, PATCOPY)
        Next
        For fY = 0 To iHeight Step mfScale
            lRet = PatBlt(picZoom.hDC, 0, Int(fY), iWidth, 1, PATCOPY)
        Next
        
        
        
        
        ' Draw the cross hair in the zoom box
        Call PatBlt(picZoom.hDC, iHeight / 2 - 9, iHeight / 2 - 9, 18, 5, PATCOPY) ' Top
        Call PatBlt(picZoom.hDC, iHeight / 2 - 9, iHeight / 2 - 9, 5, 18, PATCOPY) ' Left
        
        Call PatBlt(picZoom.hDC, iHeight / 2 - 9, iHeight / 2 + 6, 18, 5, PATCOPY) ' Bottom
        Call PatBlt(picZoom.hDC, iHeight / 2 + 6, iHeight / 2 - 9, 5, 18, PATCOPY) ' Right
        
        
        
        
        'Put the old Brush back and Delete the new one to free memory
        hBrush = SelectObject(picZoom.hDC, hOldBrush)
        lRet = DeleteObject(hBrush)
        
    End If
    
End Sub

Private Sub LoadSettings()

    'Reset mfScale
    mfScale = CSng(1000) / 100!
    
    'Force the zoom to update
    mlOldX = -100
    
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)

End Sub

Private Sub Form_Load()

    ' Move the form to the bottom right
    ' corner of the screen
    Me.Left = Screen.Width - Me.Width
    Me.Top = Screen.Height - Me.Height
    
    ' Make the picture box take up the entire form
    picZoom.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    ' Initialize the zooming functions
    Call LoadSettings

End Sub

Private Sub tmrZoom_Timer()

Dim lRet    As Long
Dim ptMouse As PointAPI

Static lElapsed As Long

    If Me.WindowState <> vbMinimized Then
        'This code runs 20 times/second*, while the form is not minimized.
        lElapsed = lElapsed + tmrZoom.Interval
        lRet = GetCursorPos(ptMouse)
        With ptMouse
            If (.X <> mlOldX) Or (.Y <> mlOldY) Or (lElapsed >= 250) Then
                'This code runs runs 4 times/second* if no mousemove,
                'or 20 times/second* when mouse is moving.
                Call DoZoom(ptMouse)
                If lElapsed >= 250 Then
                    'This code only runs 4 times/second*.
                    lRet = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
                End If
                lElapsed = 0
            End If
            mlOldX = .X
            mlOldY = .Y
        End With
    End If
    
    '* Times/second depends on processor speed. A slower processor may not
    'finish processing one timer event before the next arrives, in which
    'case the new event will be discarded.
    
End Sub


