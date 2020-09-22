VERSION 5.00
Begin VB.Form frmColorSpy 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Spy Example"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2805
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColorSpy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timSpy 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   720
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   255
      Left            =   75
      TabIndex        =   14
      Top             =   970
      Width           =   615
   End
   Begin VB.Frame fraColorInfo 
      Height          =   1175
      Left            =   750
      TabIndex        =   3
      Top             =   60
      Width           =   1940
      Begin VB.TextBox txtBlue 
         Height          =   255
         Left            =   1485
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtGreen 
         Height          =   255
         Left            =   885
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtRed 
         Height          =   255
         Left            =   280
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtHTML 
         Height          =   315
         Left            =   580
         TabIndex        =   7
         Top             =   480
         Width           =   1275
      End
      Begin VB.TextBox txtColor 
         Height          =   315
         Left            =   580
         TabIndex        =   5
         Top             =   160
         Width           =   1275
      End
      Begin VB.Label lblBlue 
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
         Height          =   180
         Left            =   1320
         TabIndex        =   12
         Top             =   850
         Width           =   135
      End
      Begin VB.Label lblGreen 
         BackStyle       =   0  'Transparent
         Caption         =   "G:"
         Height          =   180
         Left            =   720
         TabIndex        =   10
         Top             =   850
         Width           =   135
      End
      Begin VB.Label lblRed 
         BackStyle       =   0  'Transparent
         Caption         =   "R:"
         Height          =   185
         Left            =   120
         TabIndex        =   8
         Top             =   850
         Width           =   135
      End
      Begin VB.Label lblHTML 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "HTML:"
         Height          =   255
         Left            =   75
         TabIndex        =   6
         Top             =   525
         Width           =   495
      End
      Begin VB.Label lblColor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         Height          =   255
         Left            =   75
         TabIndex        =   4
         Top             =   200
         Width           =   495
      End
   End
   Begin VB.PictureBox picCursor 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   240
      Picture         =   "frmColorSpy.frx":058A
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   670
      Width           =   225
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   160
      Width           =   495
   End
   Begin VB.Frame fraBorder 
      Height          =   1445
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   2800
   End
End
Attribute VB_Name = "frmColorSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Color Spy Example
' By Chetan Sarva (csarva@ic.sunysb.edu), May 2000
' - Original by Plastik (plastik@violat0r.net)
' - Zoom by Rocky Clark

Option Explicit

'declares
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---

'conts
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const WM_NCACTIVATE = &H86

'types
Private Type PointAPI
        X As Long
        Y As Long
End Type

Private Sub cmdAbout_Click()

MsgBox "Color Spy" & vbCrLf & vbCrLf & "This program is based entirely Plastik's color spy exam and " & vbCrLf & "Rocky Clark's screen-zoom example with only a few modifi-" & vbCrLf & "cations made by Chetan Sarva to integrate the two."

End Sub

Private Sub Form_Load()

Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Sub

Private Sub picCursor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.MousePointer = 99 'set the form to allow other mouse cursors
    Me.MouseIcon = picCursor.Picture 'change mouse cursor to hold whats in the picture box
    picCursor.Visible = False 'set the visible property of the picturebox to false
    frmZoom.Show ' load the zooming window
    Call SendMessage(Me.hwnd, WM_NCACTIVATE, 1, ByVal 0&)
    timSpy.Enabled = True 'enable the timer to spy for code

End Sub

Private Sub picCursor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.MousePointer = 0 'set cursor to default
    picCursor.Visible = True 'show picture box
    Unload frmZoom ' unload the zooming window
    timSpy.Enabled = False 'disable timer to spy
    
    Clipboard.SetText txtHTML.Text ' Copy the HTML color code to the clipboard

End Sub

Private Sub timSpy_Timer()
Dim DeskTopWindow As Long, DeskTopDC As Long
Dim CurPos As PointAPI, ScreenPixel As Long
Dim strRed As String, strGreen As String
Dim strBlue As String, htmlformat As String

'use the getcursorpos api function to retrieve the current
'position of the cursor on the screen and set it to curpos
Call GetCursorPos(CurPos)

'this sets the desktop's dc in the DeskTopDC variable
DeskTopDC = GetDC(0)

'set the current pixel color in the ScreenPixel variable
'use GetPixel api function to retrieve colors from pixels
'and you use the DeskTopDC as the dc for it and we set
'the CurPos variable to hold the values of the position
'on the screen in pixels
ScreenPixel = GetPixel(DeskTopDC, CurPos.X, CurPos.Y)

'if the pictures backcolor doesn't allready = the currentolor
'pixel then dont add it to the picture box's backc
If picColor.BackColor <> ScreenPixel Then
 picColor.BackColor = ScreenPixel
End If

'set the txtColor's text to the background color of the pixel
txtColor.Text = picColor.BackColor

'strRed$ gets set with the Red RGB value
strRed$ = CStr(picColor.BackColor And 255)
'strGreen$ gets set with the Green RGB value
strGreen$ = CStr(picColor.BackColor \ 256 And 255)
'strBlue$ gets set with the Blue RGB value
strBlue$ = CStr(picColor.BackColor \ 65536 And 255)

'set the RGB values in the RGB textbox's
txtRed.Text = strRed$
txtGreen.Text = strGreen$
txtBlue.Text = strBlue$

' Convert the RGB values to HTML Hex format
strRed = Hex(strRed): If strRed = "0" Then strRed = "00"
strGreen = Hex(strGreen): If strGreen = "0" Then strGreen = "00"
strBlue = Hex(strBlue): If strBlue = "0" Then strBlue = "00"

' Put the HTML color string into the textbox
txtHTML.Text = "#" & strRed & strGreen & strBlue

End Sub
