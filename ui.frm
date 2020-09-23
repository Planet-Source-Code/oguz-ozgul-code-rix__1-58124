VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   552
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer15 
      Interval        =   40
      Left            =   1680
      Top             =   4080
   End
   Begin VB.PictureBox picCommand 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1455
      Left            =   4680
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   6
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Timer Timer14 
      Interval        =   1000
      Left            =   1680
      Top             =   3480
   End
   Begin VB.Timer Timer13 
      Interval        =   10
      Left            =   1680
      Top             =   2880
   End
   Begin VB.PictureBox picGame 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00008000&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   4
      Top             =   6600
      Width           =   4335
   End
   Begin VB.PictureBox picVB 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   3
      Top             =   4560
      Width           =   4335
   End
   Begin VB.Timer Timer12 
      Interval        =   100
      Left            =   2880
      Top             =   2280
   End
   Begin VB.Timer Timer11 
      Interval        =   50
      Left            =   2280
      Top             =   2280
   End
   Begin VB.Timer Timer10 
      Interval        =   39
      Left            =   4080
      Top             =   1680
   End
   Begin VB.Timer Timer9 
      Interval        =   45
      Left            =   4680
      Top             =   1680
   End
   Begin VB.Timer Timer8 
      Interval        =   51
      Left            =   5280
      Top             =   1680
   End
   Begin VB.Timer Timer7 
      Interval        =   57
      Left            =   5880
      Top             =   1680
   End
   Begin VB.Timer Timer6 
      Interval        =   10
      Left            =   1680
      Top             =   2280
   End
   Begin VB.Timer Timer5 
      Interval        =   33
      Left            =   3480
      Top             =   1680
   End
   Begin VB.Timer Timer4 
      Interval        =   27
      Left            =   2880
      Top             =   1680
   End
   Begin VB.Timer Timer3 
      Interval        =   21
      Left            =   2280
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Interval        =   15
      Left            =   1680
      Top             =   1680
   End
   Begin VB.PictureBox digits 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   6000
      Picture         =   "ui.frx":0000
      ScaleHeight     =   750
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1680
      Top             =   1080
   End
   Begin VB.PictureBox picASM 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.PictureBox picC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   2
      Top             =   2400
      Width           =   4335
   End
   Begin VB.PictureBox picWatch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FontTransparent =   0   'False
      ForeColor       =   &H00008000&
      Height          =   3735
      Left            =   4560
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   239
      TabIndex        =   5
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


Private Const W_OF_CHAR As Long = 15
Private Const H_OF_CHAR As Long = 18
Private ffASM As Long
Private ffC As Long
Private ffVB As Long

Private Type tetrixPos
    filled As Boolean
    color As Long
End Type

Private Type tetrixPiece
    x As Long
    y As Long
    color As Long
End Type

Dim colors(3) As Long

Private msgString As String

Private Sub Form_Click()
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    End
End Sub


Private Sub Form_Load()
    Dim nV As Long
    nV = Int(((Screen.Height / Screen.TwipsPerPixelY) / H_OF_CHAR) / 3)
    nV = 10
    Move 0, 0, Screen.Width, Screen.Height
    
    
    picASM.Move 0, 0, W_OF_CHAR * 15 - 1, H_OF_CHAR * nV - 1
    picC.Move 0, H_OF_CHAR * (nV + 1), W_OF_CHAR * 15 - 1, H_OF_CHAR * nV - 1
    picVB.Move 0, H_OF_CHAR * (2 * nV + 2), W_OF_CHAR * 15 - 1, H_OF_CHAR * nV - 1
    picGame.Move 0, H_OF_CHAR * (3 * nV + 3), 3 * (Screen.Width / Screen.TwipsPerPixelX) / W_OF_CHAR, 3 * (Screen.Height / Screen.TwipsPerPixelY) / H_OF_CHAR
    picWatch.Move (Screen.Width / Screen.TwipsPerPixelX) - 196, 0, 196, 388
    picCommand.Move (Screen.Width / Screen.TwipsPerPixelX) - 196, 392 + H_OF_CHAR, 196, 300
    
    ffASM = FreeFile
    Open App.Path & "\codeASM.txt" For Input Access Read As ffASM
    ffC = FreeFile
    Open App.Path & "\codeC.txt" For Input Access Read As ffC
    ffVB = FreeFile
    Open App.Path & "\codeVB.txt" For Input Access Read As ffVB
    
    msgString = _
                " (C) BY OGUZ OZGUL, 2004" & vbCrLf & _
                " CODE MATRIX SCREEN SAVER V1.0" & vbCrLf & _
                " PRESS THE MOUSE BUTTON TO EXIT" & vbCrLf & _
                vbCrLf & _
                " VISUAL BASIC IS BACK" & vbCrLf & _
                " ALL THE THREADS RUNNING ON THE SCREEN" & vbCrLf & _
                " ARE CONTROLLED BY A TIMER" & vbCrLf & _
                vbCrLf & _
                " THE VB TIMER OBJECT IS POWERFUL" & vbCrLf & _
                " AND IS ABLE TO GIVE THE FEELING" & vbCrLf & _
                " OF A MULTI THREADED APPLICATION" & vbCrLf & _
                vbCrLf & _
                " DON'T FORGET TO GO BACK TO THE" & vbCrLf & _
                " REAL LIFE" & vbCrLf & _
                vbCrLf & _
                " ENJOY :)"
                 
    colors(0) = RGB(0, 32, 0)
    colors(1) = RGB(0, 64, 0)
    colors(2) = RGB(0, 96, 0)
    colors(3) = RGB(0, 128, 0)
                
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Close ffASM
    Close ffC
    Close ffVB
    Close
End Sub

Private Sub picASM_Click()
    End
End Sub

Private Sub Timer1_Timer()
    Static started As Boolean
    Dim i As Long
    If Not started Then
        picVB.Visible = False
        picC.Visible = False
        picASM.Visible = False
        picWatch.Visible = False
        picGame.Visible = False
        picCommand.Visible = False
        For i = 0 To 3000
            printNumber Int(Rnd * 999999999)
        Next i
        started = True
        picVB.Visible = True
        picC.Visible = True
        picASM.Visible = True
        picWatch.Visible = True
        picGame.Visible = True
        picCommand.Visible = True
    End If
    Randomize Timer
    printNumber Int(Rnd * 999999999)
End Sub

Private Sub printNumber(ByVal number As Long)
    Dim i As Long, j As Long
    Dim x As Long
    Dim y As Long
    Dim s As String
    Dim cY As Long
    s = Right("000000000" & number, 9)
    Randomize Timer
    For i = 1 To Len(s)
        x = getRandomPos
        y = Int(Rnd * ((Form1.Height / Screen.TwipsPerPixelY) / H_OF_CHAR)) * H_OF_CHAR
        'If Rnd() < 0.1 Then cY = 23 Else cY = 0
        cY = 0
        BitBlt hDC, x + i * W_OF_CHAR, y, W_OF_CHAR, H_OF_CHAR, digits.hDC, Mid(s, i, 1) * W_OF_CHAR, cY, SRCCOPY
    Next i
    'BitBlt hDC, 10, 33, 165, 800, hDC, 10, 10, SRCCOPY

End Sub


Private Sub Timer11_Timer()
    Static lblCount As Long
    Static clear As Boolean
    HandleCodeWindow "C++ WINDOW", picC, ffC, lblCount, Timer11, clear, 500, 25
End Sub

Private Sub Timer12_Timer()
    Static lblCount As Long
    Static clear As Boolean
    HandleCodeWindow "VISUAL BASIC WINDOW", picVB, ffVB, lblCount, Timer12, clear, 1000, 50
End Sub

Private Sub Timer13_Timer()
    Dim s As String
    Dim i As Long
    Dim j As Long
    'Randomize Timer
        
    'picGame.PSet (Int(Rnd * picGame.Width), Int(Rnd * picGame.Height)), vbGreen

End Sub

Private Sub Timer14_Timer()

    Static anim As Long
    Static xb As Long
    Static yb As Long
    Static color As Long
    Static posFill(12, 24) As tetrixPos
    Static pieceDropping As Boolean
    Static flyingPiece As tetrixPiece
    Static gameOver As Long
    Dim cColor As Long
    Dim i As Long, j As Long
    
    Select Case anim
        Case 1
            picWatch.Line (0, 0)-(0, picWatch.Height), RGB(0, 64, 0)
            gameOver = False
            For i = 0 To 12
                For j = 0 To 24
                    posFill(i, j).filled = False
                    posFill(i, j).color = -1
                Next j
            Next i
        Case 2
            picWatch.Line (0, 0)-(0, picWatch.Height), vbGreen
            picWatch.Line (0, 0)-(picWatch.Width, 0), RGB(0, 64, 0)
        Case 3
            picWatch.Line (0, 0)-(picWatch.Width, 0), vbGreen
            picWatch.Line (picWatch.Width - 3, 0)-(picWatch.Width - 3, picWatch.Height - 3), RGB(0, 64, 0)
        Case 4
            picWatch.Line (picWatch.Width - 3, 0)-(picWatch.Width - 3, picWatch.Height - 3), RGB(0, 128, 0)
            picWatch.Line (0, picWatch.Height - 3)-(picWatch.Width - 3, picWatch.Height - 3), RGB(0, 64, 0)
        Case 5, 7, 9
            If anim = 5 Then Timer14.interval = 100
            If anim = 9 Then Timer14.interval = 10
            picWatch.Line (0, 0)-(picWatch.Width - 3, picWatch.Height - 3), RGB(0, 128, 0), B
        Case 6, 8
            picWatch.Line (0, 0)-(picWatch.Width - 3, picWatch.Height - 3), RGB(0, 64, 0), B
        Case 10 To 201
            picWatch.Line (anim - 9, 1)-(anim - 9, picWatch.Height - 3), RGB(0, 8, 0)
            picWatch.Line (anim - 8, 1)-(anim - 8, picWatch.Height - 3), RGB(0, 128, 0)
        Case 202 To 205
            Timer14.interval = 50
        Case 206 To 217
            picWatch.Line (1, (anim - 205) * 16)-(picWatch.Width - 3, (anim - 205) * 16), RGB(0, 32, 0)
            picWatch.Line (1, (anim - 194) * 16)-(picWatch.Width - 3, (anim - 194) * 16), RGB(0, 32, 0)
            picWatch.Line ((anim - 205) * 16, 1)-((anim - 205) * 16, picWatch.Height - 3), RGB(0, 32, 0)
            color = 8
        Case Is > 217
            Timer14.interval = 10
            If Not pieceDropping Then
                ' Create a Piece
                flyingPiece.x = Int(Rnd * 12)
                flyingPiece.y = 0
                flyingPiece.color = Int(Rnd * 3.99)
                pieceDropping = True
            Else
                ' Drop piece
                picWatch.Line (flyingPiece.x * 16 + 1, flyingPiece.y * 16 + 1)-(flyingPiece.x * 16 + 15, flyingPiece.y * 16 + 15), RGB(0, 8, 0), BF
                flyingPiece.y = flyingPiece.y + 1
            End If
            
            If posFill(flyingPiece.x, flyingPiece.y + 1).filled = True Or flyingPiece.y = 23 Then
                ' Stop piece
                pieceDropping = False
                posFill(flyingPiece.x, flyingPiece.y).filled = True
                posFill(flyingPiece.x, flyingPiece.y).color = flyingPiece.color
                
                If posFill(flyingPiece.x, flyingPiece.y).color <> posFill(flyingPiece.x, flyingPiece.y + 1).color Then
                    ' Draw the piece
                    picWatch.Line (flyingPiece.x * 16 + 1, flyingPiece.y * 16 + 1)-(flyingPiece.x * 16 + 15, flyingPiece.y * 16 + 15), colors(flyingPiece.color), BF
                Else
                    ' Clear two pieces
                    picWatch.Line (flyingPiece.x * 16 + 1, flyingPiece.y * 16 + 1)-(flyingPiece.x * 16 + 15, flyingPiece.y * 16 + 15), RGB(0, 8, 0), BF
                    picWatch.Line (flyingPiece.x * 16 + 1, (flyingPiece.y + 1) * 16 + 1)-(flyingPiece.x * 16 + 15, (flyingPiece.y + 1) * 16 + 15), RGB(0, 8, 0), BF
                    With posFill(flyingPiece.x, flyingPiece.y)
                        .filled = False
                        .color = -1
                    End With
                    With posFill(flyingPiece.x, flyingPiece.y + 1)
                        .filled = False
                        .color = -1
                    End With
                    
                End If
            Else
                ' Draw the piece
                picWatch.Line (flyingPiece.x * 16 + 1, flyingPiece.y * 16 + 1)-(flyingPiece.x * 16 + 15, flyingPiece.y * 16 + 15), colors(flyingPiece.color), BF
            End If
            If flyingPiece.y = 0 And Not pieceDropping Then
                gameOver = True
            End If
            
        Case Else
            ' First trip
            Timer14.interval = 125
    End Select
    If Not gameOver Then
        anim = anim + 1
    Else
        'resart
        anim = 1
    End If
    
End Sub

Private Sub Timer15_Timer()
    Dim x As Long, y As Long
    Static lastX As Long, lastY As Long
    Static show As Boolean
    Static pos As Long
    Static clearDone As Boolean
    Static clearY As Long
    
    If pos = 0 Then
        picCommand.Line (0, 0)-(picCommand.Width - 3, picCommand.Height - 3), RGB(0, 128, 0), B
        picCommand.CurrentX = 0
        picCommand.CurrentY = 0
        picCommand.Print ""
        clearDone = False
        clearY = 1
    End If
    
    If pos = 0 Then pos = 1
    
    If pos < Len(msgString) + 1 Then
        picCommand.Print Mid(msgString, pos, 1);
    End If
    
    pos = pos + 1
    
    'picCommand.Line (0, 0)-(picCommand.Width - 3, picCommand.Height - 3), RGB(pos, pos, pos), B
    
    If pos > Len(msgString) + 25 Then
        If (Not clearDone) And (clearY < picCommand.Height - 6) Then
            picCommand.Line (1, 1)-(picCommand.Width - 4, clearY), vbBlack, BF
            clearY = clearY + 5
            If pos > Len(msgString) + 100 Then
                clearDone = True
            End If
        Else
            pos = 0
            'picCommand.Cls
        End If
    End If
    
End Sub

Private Sub Timer2_Timer()
    Static x As Long
    Static y As Long
    Static xOld As Long
    timerHandler x, y, xOld
End Sub

Private Sub Timer3_Timer()
    Static x As Long
    Static y As Long
    Static xOld As Long
    timerHandler x, y, xOld
End Sub

Private Sub Timer4_Timer()
    Static x As Long
    Static y As Long
    Static xOld As Long
    timerHandler x, y, xOld
End Sub

Private Sub Timer5_Timer()
    Static x As Long
    Static y As Long
    Static xOld As Long
    timerHandler x, y, xOld
End Sub
Private Sub Timer7_Timer()
    Static x As Long
    Static y As Long
    Static xOld As Long
    timerHandler x, y, xOld
End Sub

Private Sub Timer8_Timer()
    Static x As Long
    Static y As Long
    Static xOld As Long
    timerHandler x, y, xOld
End Sub

Private Sub Timer9_Timer()
    Static x As Long
    Static y As Long
    Static xOld As Long
    timerHandler x, y, xOld
End Sub

Private Sub Timer10_Timer()
    Static x As Long
    Static y As Long
    Static xOld As Long
    timerHandler x, y, xOld
End Sub


Private Sub Timer6_Timer()
    Static lblCount As Long
    Static clear As Boolean
    HandleCodeWindow "ASSMEBLY WINDOW", picASM, ffASM, lblCount, Timer6, clear, 400, 15
End Sub

Private Sub timerHandler(ByRef x As Long, ByRef y As Long, ByRef xOld As Long)
    'Form1.CurrentX = 0
    'Form1.CurrentY = 0
    'Form1.Print "timerHandler() called with parameters " & x & ", " & y & ", " & xOld
    If x = 0 And xOld = 0 Then
        x = getRandomPos
    End If
    BitBlt hDC, x, y - H_OF_CHAR, W_OF_CHAR, H_OF_CHAR, digits.hDC, Int(Rnd * 10) * W_OF_CHAR, H_OF_CHAR, SRCCOPY
    y = y + H_OF_CHAR
    If y > Form1.Height / Screen.TwipsPerPixelY + 2 * Form1.TextHeight("a") Then
        y = 0
        xOld = x
        x = getRandomPos
    End If
    
    BitBlt hDC, x, y, W_OF_CHAR, H_OF_CHAR, digits.hDC, Int(Rnd * 10) * W_OF_CHAR, 2 * H_OF_CHAR, SRCCOPY
    BitBlt hDC, xOld, y, W_OF_CHAR, H_OF_CHAR, digits.hDC, Int(Rnd * 10) * W_OF_CHAR, 0, SRCCOPY
    
    Dim gX As Long, gY As Long
    gX = x * (picGame.Width / (Screen.Width / Screen.TwipsPerPixelX))
    gY = y * (picGame.Height / (Screen.Height / Screen.TwipsPerPixelY))
    picGame.Line (gX, gY)-(gX + 2, gY + 2), vbGreen, BF
    picGame.Line (gX, gY - 3)-(gX + 2, gY - 1), RGB(0, 64, 0), BF
    
    gX = xOld * (picGame.Width / (Screen.Width / Screen.TwipsPerPixelX))
    picGame.Line (gX, gY)-(gX + 2, gY + 2), vbBlack, BF

End Sub

Private Function getRandomPos()
    Randomize Timer
    getRandomPos = Int(Rnd * (((Form1.Width / Screen.TwipsPerPixelX)) / W_OF_CHAR)) * W_OF_CHAR
End Function

Private Function HandleCodeWindow(ByVal codeWindowName As String, _
                                  ByRef picB As PictureBox, _
                                  ByVal ff As Long, _
                                  ByRef lblCount As Long, _
                                  ByRef tmr As Timer, _
                                  ByRef clear As Boolean, _
                                  ByVal wait As Long, _
                                  ByVal interval As Long)
    
    Dim s As String
    
    If EOF(ff) Then
        Seek ff, 1
    End If
    Line Input #ff, s
    
    'lastY = picB.CurrentY
    
    'If picB.CurrentY > picB.Height - 2 * picB.TextHeight("a") Then
    'End If
    If lblCount = 0 Then
        picB.Line (0, 0)-(picB.Width - 3, picB.Height - 3), RGB(0, 128, 0), B
        picB.CurrentX = 0
        picB.CurrentY = 5
        picB.FontBold = True
        picB.Print " " & codeWindowName
        picB.FontBold = False
        picB.CurrentY = picB.CurrentY + 5
        'picB.Print String(240, "~")
        'picB.CurrentY = picB.Height - 1 * picB.TextHeight("a")
        'picB.Print String(240, "~")
        'picB.CurrentY = picB.TextHeight("a") * 2
        lblCount = 1
    End If
    
    If picB.CurrentY > picB.Height - picB.TextHeight("a") * 3 Then
        If clear Then
            picB.Cls
            picB.Line (0, 0)-(picB.Width - 3, picB.Height - 3), RGB(0, 128, 0), B
            picB.CurrentX = 0
            picB.CurrentY = 5
            picB.FontBold = True
            picB.Print " " & codeWindowName
            picB.FontBold = False
            picB.CurrentY = picB.CurrentY + 5
            'picB.Print String(240, "~")
            'picB.CurrentY = picB.Height - 1 * picB.TextHeight("a")
            'picB.Print String(240, "~")
            'picB.CurrentY = picB.TextHeight("a") * 2
            clear = False
            tmr.interval = interval
        Else
            tmr.interval = wait + Rnd * wait * 2
            clear = True
        End If
    End If
    
    picB.Print " " & s & Space(100)

End Function
