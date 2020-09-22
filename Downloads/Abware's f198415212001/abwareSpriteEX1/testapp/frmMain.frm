VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr1 
      Interval        =   10
      Left            =   1410
      Top             =   1140
   End
   Begin VB.Image img 
      Height          =   465
      Left            =   180
      Top             =   2370
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private o As ABSPRITEEX1Lib.Scene
Private m_stp           As Boolean
Private m_xx             As Integer
Private m_yy             As Integer
Private m_x()           As Integer
Private m_y()           As Integer
Private m_Speed()       As Integer
Private m_xdir()        As Integer
Private m_ydir()        As Integer
Private m_Counter       As Integer
Private m_i()           As Integer

Private m_timer             As Integer
Private m_tm                As Integer
Private m_run               As Boolean

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub Form_Click()
    m_stp = True
End Sub

Private Sub AddImages()

    On Error GoTo errdeal
    Dim objpic As StdPicture
    
    img.Picture = LoadPicture(App.Path & "\ready.bmp")
    o.AddImage "donut", App.Path & "\donut.bmp"
    o.AddImage "ready", App.Path & "\ready.bmp"
    o.AddImage "font1", App.Path & "\fonts1.bmp"
    Exit Sub
    
errdeal:
    Set objpic = LoadPicture(App.Path & "\ready.jpg")
    img.Picture = objpic
    SavePicture img.Picture, App.Path & "\ready.bmp"
    Resume
End Sub

Private Sub AddBackgrounds()
    On Error GoTo errdeal
    Dim objpic As StdPicture
    
    img.Picture = LoadPicture(App.Path & "\rockies.bmp")
    o.AddBackground "rockies", App.Path & "\rockies.bmp"
    Exit Sub
    
errdeal:
    Set objpic = LoadPicture(App.Path & "\rockies.jpg")
    img.Picture = objpic
    SavePicture img.Picture, App.Path & "\rockies.bmp"
    Resume
End Sub

Private Sub AddSprites()
    o.AddSprite "me", "ready", 2, 1   'default picture
End Sub

Private Sub AddActions()
    o.AddAction "ready", 0, 1, 1, 0
    
End Sub

Private Sub Form_Load()

    
    m_run = False
    ChDir App.Path
    Set o = New ABSPRITEEX1Lib.Scene
    
    m_stp = False
    m_Counter = 0
    
    ReDim m_x(0)
    ReDim m_y(0)
    ReDim m_xdir(0)
    ReDim m_ydir(0)
    ReDim m_Speed(0)
    
    
    Me.Show
    o.HWND_Scene = Me.hWnd
    o.Init 640, 480, 16
    o.RenderSpeed = 0
    
    'init scene objects
    AddImages
    AddBackgrounds
    AddSprites
    AddActions
    o.AddFont "font1", "font1", 12
    o.AddLabel "lblAbware", "ABWARE FAST GAME ENGINE", 100, 200
    o.AddLabel "lblAlien", "ALIENS: ", 460, 10
    o.AddLabel "lblAliensNum", "0000", 460, 100
    o.AddLabel "lbldemo", "ITS JUST A SIMPLE DEMO", 50, 200
    o.StartAction "me", "ready", 1
    o.SetActionSpeed "me", 5

    m_run = True
    
End Sub

Private Sub AddSprite()

    m_Counter = m_Counter + 1
    o.UpdateLabel "lblAliensNum", Format(m_Counter, "0000")
    ReDim Preserve m_x(m_Counter)
    ReDim Preserve m_y(m_Counter)
    ReDim Preserve m_xdir(m_Counter)
    ReDim Preserve m_ydir(m_Counter)
    ReDim Preserve m_Speed(m_Counter)
    
    m_x(m_Counter) = Rnd(1) * 600
    m_y(m_Counter) = Rnd(1) * 400
    m_xdir(m_Counter) = Rnd(1) * 1
    If m_xdir(m_Counter) = 0 Then m_xdir(m_Counter) = 1
    m_ydir(m_Counter) = Rnd(1) * 1
    If m_ydir(m_Counter) = 0 Then m_ydir(m_Counter) = 1
    m_Speed(m_Counter) = Rnd(1) * 5
    If m_Speed(m_Counter) = 0 Then m_Speed(m_Counter) = 1
    
    o.AddSprite "s" & m_Counter, "donut", 1, 1
    
End Sub

Private Sub tmr1_Timer()
    
    Dim i           As Long '
    
    If Not m_run Then Exit Sub
    
    If GetAsyncKeyState(vbKeyUp) <> 0 Then
        m_yy = m_yy - 5
    End If
    If GetAsyncKeyState(vbKeyDown) <> 0 Then
        m_yy = m_yy + 5
    End If
    If GetAsyncKeyState(vbKeyLeft) <> 0 Then
        m_xx = m_xx - 8
    End If
    If GetAsyncKeyState(vbKeyRight) <> 0 Then
        m_xx = m_xx + 8
    End If
    If GetAsyncKeyState(vbKeyEscape) <> 0 Then
        m_run = False
        o.Restore
    End If

    
    For i = 1 To m_Counter
        m_x(i) = m_x(i) + m_Speed(i) * m_xdir(i)
        m_y(i) = m_y(i) + m_Speed(i) * m_ydir(i)
        If m_x(i) < 0 Or m_x(i) > 600 Then m_xdir(i) = m_xdir(i) * -1
        If m_y(i) < 0 Or m_y(i) > 400 Then m_ydir(i) = m_ydir(i) * -1
        o.MoveSprite "s" & i, m_x(i), m_y(i), 0
    Next i
    
    m_timer = m_timer + 1
    If m_timer > 10 Then
        AddSprite
        m_timer = 0
    End If
    
    o.MoveSprite "me", m_xx, m_yy, 0
    
    If m_Counter >= 1 Then
        For i = 1 To m_Counter
            If o.IsColide("me", "s" & i) = 1 Then
                'GoTo ext
            End If
        Next i
    End If
    DoEvents
    o.Render "rockies"

End Sub
