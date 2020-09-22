VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SpyHunter"
   ClientHeight    =   6165
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   7155
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTarget 
      AutoRedraw      =   -1  'True
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6075
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnureset 
         Caption         =   "&Reset"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuControls 
      Caption         =   "Controls"
      Begin VB.Menu mnuInfo 
         Caption         =   "&Info"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim bored As Boolean

Dim mx As Single
Dim my As Single
Dim mz As Single

Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim carx As Long

Dim keyLeft As Boolean
Dim keyRight As Boolean
Dim keyFire As Boolean
Dim keyAccel As Boolean
Dim keyBrake As Boolean
Dim keyJump As Boolean
Dim pause As Boolean
Dim keyStart As Boolean
Dim level As Long

Dim Npos As Long
Dim Frameno As Long
Dim expframe As Single
Dim dead As Boolean
Dim Gas As Single
Dim Score As Long
Dim maxTime As Single
Dim lastPressed As Boolean
Dim jumpPressed As Boolean
Dim deadtimer As Long
Dim nextlevel As Boolean

Dim EnemyCol As New Collection

Dim RoadMap(0 To 15, 0 To 299) As Byte
Dim ObjectMap(0 To 15, 0 To 299) As Byte
Dim RoadInfo(0 To 54) As Byte

Const mPort = &H378

Private Sub Form_Load()
    picTarget.width = 512 * Screen.TwipsPerPixelX
    picTarget.Height = 384 * Screen.TwipsPerPixelY
    Me.width = picTarget.width + 60
    Me.Height = picTarget.Height + 650
    
    picTarget.ScaleMode = vbPixels
    level = 1
    
   
    Open "roadinfo.dat" For Binary As 1
    Get #1, , RoadInfo
    Close 1
    
    InitDX frmMain
    
    With ddsd
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
    
    Set ddsPrimary = DDraw.CreateSurface(ddsd)
    
    Set DDClip = DDraw.CreateClipper(0)
    DDClip.SetHWnd picTarget.hWnd
    ddsPrimary.SetClipper DDClip
    
    With ddsd
        .lFlags = DDSD_CAPS + DDSD_WIDTH + DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 384
        .lWidth = 512
    End With
    
    Set ddsBG = DDraw.CreateSurfaceFromFile(App.Path & "\title.bmp", ddsd)
    
    With ddsd
        .lFlags = DDSD_CAPS + DDSD_WIDTH + DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 160
        .lWidth = 352
    End With
    
    Set ddsRoad = DDraw.CreateSurfaceFromFile(App.Path & "\road.bmp", ddsd)
    
    Dim ddc As DDCOLORKEY
    Dim rRect As RECT
    Dim col As Long
    
    ddsRoad.Lock rRect, ddsd, DDLOCK_WAIT, 0
    col = ddsRoad.GetLockedPixel(0, 0)
    ddsRoad.Unlock rRect
    
    ddc.high = col
    ddc.low = col
    
    ddsRoad.SetColorKey DDCKEY_SRCBLT, ddc
    
    With ddsd
        .lFlags = DDSD_CAPS + DDSD_WIDTH + DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 160
        .lWidth = 352
    End With
    
    Set ddsObjects = DDraw.CreateSurfaceFromFile(App.Path & "\objectdat.bmp", ddsd)
    
    ddsObjects.Lock rRect, ddsd, DDLOCK_WAIT, 0
    col = ddsObjects.GetLockedPixel(0, 0)
    ddsObjects.Unlock rRect
    
    ddc.high = col
    ddc.low = col
    
    ddsObjects.SetColorKey DDCKEY_SRCBLT, ddc
    
    With ddsd
        .lFlags = DDSD_CAPS + DDSD_WIDTH + DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 160
        .lWidth = 352
    End With
    
    Set ddsOther = DDraw.CreateSurfaceFromFile(App.Path & "\other.bmp", ddsd)
    
    ddsOther.Lock rRect, ddsd, DDLOCK_WAIT, 0
    col = ddsOther.GetLockedPixel(0, 0)
    ddsOther.Unlock rRect
    
    ddc.high = col
    ddc.low = col
    
    ddsOther.SetColorKey DDCKEY_SRCBLT, ddc
    
    With ddsd
        .lFlags = DDSD_CAPS + DDSD_WIDTH + DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 448
        .lWidth = 576
    End With
    
    Set ddsBuffer = DDraw.CreateSurface(ddsd)
    
    Dim mrect As RECT
    mrect.Bottom = 448
    mrect.Right = 576
    
    ddsBuffer.BltColorFill mrect, 0
    
    Me.Show
    
    mnureset_Click
    
    'main game loop
    While Not bored
    
        DoEvents
        
        'game plays as long as not dead, not pause, or not moving to the next level
        If Not dead And Not pause And Not nextlevel Then
        
            'set movement speed
            'dx = IIf(dy > 1, IIf(keyLeft, dx - 0.2, IIf(keyRight, dx + 0.2, dx - (Sgn(dx)) * 0.2)), 0) '* (dy / 6)
            dx = IIf(dy > 1, IIf(keyLeft, dx - 0.2, IIf(keyRight, dx + 0.2, IIf(Abs(dx) > 0.01, dx - (Sgn(dx)) * 0.2, 0))), 0)    '* (dy / 6)
            dy = IIf(keyAccel, dy + 0.05, dy - 0.01)
            
            'react to brakes
            If keyBrake And (mz = 0) Then
                dy = dy - 0.2
            End If
            
            'react to jump
            If keyJump Then
                If Not jumpPressed Then
                    jumpPressed = True
                    If (mz = 0) Then
                        dz = 1
                        Gas = Gas - 2
                    End If
                End If
            Else
                jumpPressed = False
            End If
            
            If keyFire Then
                savey = Npos:   savex = xpos
                temp1 = ObjectMap(savex, savey + 1)
                ObjectMap(savex, savey + 1) = 30
                temp2 = ObjectMap(savex, savey + 2)
                ObjectMap(savex, savey + 2) = 30
                temp3 = ObjectMap(savex, savey + 3)
                ObjectMap(savex, savey + 3) = 31
            End If
            
            xpos = ((mx + 16) \ 32)
            If xpos > 15 Then xpos = 15
            
            'react to objects on screen

            If (RoadInfo(RoadMap(xpos, Npos)) And 1) = 0 And mz = 0 Then
                dead = True
            End If
            
'            If ObjectMap(xpos, Npos) = 21 And mz = 0 Then
'                dead = True
'            End If
            
            If ObjectMap(xpos, Npos) = 54 Then
                nextlevel = True
            End If
            
            If ObjectMap(xpos, Npos) = 5 And mz = 0 Then
                ObjectMap(xpos, Npos) = 0
                Gas = Gas + 10
                If Gas > 100 Then Gas = 100
            End If
            
            If (((RoadInfo(RoadMap(xpos, Npos)) \ 2) And 1) = 1) And (mz = 0) Then
                dy = 2
            End If
            
            If (((RoadInfo(RoadMap(xpos, Npos)) \ 4) And 1) = 1) And (mz = 0) And (dy > 0) Then
                dz = 1
            End If
            
            If keyAccel Then
                Gas = Gas - 0.05
                If dy > 5 Then Score = Score + 1
            End If
            
            If CLng(Gas) < 32 Then dead = True
            
            maxTime = maxTime - 0.02
            If CLng(maxTime) < 0 Then dead = True
            
        Else
            If dead Then
                dy = dy - 0.2
                dx = dx - (Sgn(dx)) * 0.6
            End If
        End If
            
        If keyStart And Not dead Then
            If Not lastPressed Then
                lastPressed = True
                pause = Not pause
            End If
        Else
            lastPressed = False
        End If
            
        If dead Then
            If deadtimer = 0 Then
                deadtimer = 300
            Else
                deadtimer = deadtimer - 1
                If deadtimer = 0 Then mnureset_Click
            End If
        End If
            
        If nextlevel Then
            dy = 0
            my = my - 10
            If my < -32 Then
                If deadtimer = 0 Then
                    deadtimer = 300
                Else
                    deadtimer = deadtimer - 1
                    If deadtimer = 0 Then newLevel
                End If
            End If
        End If
            
        If keyStart And dead Then
            mnureset_Click
        End If
            
        If Not pause Then
            If Abs(dx) > 5 Then dx = 5 * Sgn(dx)
            
            If dy > 7 Then dy = 7
            If dy < 0 Then dy = 0
            
            dz = dz - 0.05
            
            If mz < 0 Then
                dz = 0
                mz = 0
            End If
            
            mx = mx + dx
            mz = mz + dz
            If mx < 0 Then mx = 0
            If mx > 512 Then mx = 512
            
            Frameno = Frameno + dy
            
            If Frameno > 32 Then
                Frameno = 0
                Npos = Npos + 1
                If Npos > 288 Then Npos = 0
            End If
        End If
        
                
        'after all that, we must RENDER
        Render mx, my, mz
        ObjectMap(savex, savey + 1) = temp1
        ObjectMap(savex, savey + 2) = temp2
        ObjectMap(savex, savey + 3) = temp3
        
        'Slow down to a playable speed
        Sleep 5
    Wend
    
    DestroyDX
    
    End
End Sub

Public Sub Render(x As Single, y As Single, Z As Single)
    Dim dRect As RECT
    Dim sRect As RECT
    Dim wRect As RECT
    Dim tile As Long
    Dim rx As Long
    Dim ry As Long
    
    DX7.GetWindowRect picTarget.hWnd, wRect
    
    For ry = 11 To 0 Step -1
        For rx = 0 To 15
        
            DrawRoad CLng(RoadMap(rx, Npos + ry)), rx * 32, (11 - ry) * 32 + Frameno
            
            tile = ObjectMap(rx, Npos + ry)
            
            Select Case tile
            Case 1
            Case 2, 3, 4
                ObjectMap(rx, Npos + ry) = 0
                Dim newEnemy As New clsEnemy
                newEnemy.mtype = tile
                newEnemy.prevTile = 0
                newEnemy.x = rx
                newEnemy.y = ry + Npos
                newEnemy.dy = 0.02 + Rnd(1) * 0.2
                n = Rnd(1)
                newEnemy.Tag = "Enemy" & n
                EnemyCol.Add newEnemy, "Enemy" & n
                Set newEnemy = Nothing
            Case Else
                DrawObject tile, rx * 32, (11 - ry) * 32 + Frameno
            End Select
        
        
        Next
    Next
    
    Dim Enemy As clsEnemy
    
    For Each Enemy In EnemyCol
    
        If Not Enemy.dead And Not pause Then
        ObjectMap(Enemy.x, Enemy.y) = Enemy.prevTile
        
        If (RoadInfo(RoadMap(Enemy.x, Enemy.y)) And 1) = 0 Then
            Enemy.dead = True
        End If
        
        If Enemy.prevTile = 31 Then
            Enemy.dead = True
        End If
        
        'look ahead
        epos = Enemy.y + 2
        If epos < 299 Then
            Select Case RoadMap(Enemy.x, epos)
            Case 4, 5, 26, 27
                Enemy.dx = Enemy.dx + Enemy.dy / 20
            Case 6, 7, 28, 29
                Enemy.dx = Enemy.dx - Enemy.dy / 20
            Case Else
                Enemy.dx = Enemy.dx / 2
            End Select
            End If
            
            Enemy.x = Enemy.x + Enemy.dx
            Enemy.y = Enemy.y + Enemy.dy
            
            If Enemy.y < Npos - 11 Then EnemyCol.Remove Enemy.Tag
            If Enemy.y > Npos + 22 Then EnemyCol.Remove Enemy.Tag
            If Enemy.y > 299 Then EnemyCol.Remove Enemy.Tag
            
            Enemy.prevTile = ObjectMap(Enemy.x, Enemy.y)
            ObjectMap(Enemy.x, Enemy.y) = 21
        End If
        
        If Not Enemy.dead Then
            Select Case Enemy.mtype
            Case 4
                'truck type?
                DrawObject Enemy.mtype, Enemy.x * 32, (11 - Enemy.y + Npos) * 32 + Frameno
                DrawObject Enemy.mtype + 11, Enemy.x * 32, (11 - Enemy.y + Npos + 1) * 32 + Frameno
                DrawObject Enemy.mtype + 22, Enemy.x * 32, (11 - Enemy.y + Npos + 2) * 32 + Frameno
            Case Else
                DrawObject 11, Enemy.x * 32, (11 - Enemy.y + Npos) * 32 + Frameno
                DrawObject Enemy.mtype, Enemy.x * 32, (11 - Enemy.y + Npos) * 32 + Frameno
            End Select
        Else
            DrawOther CLng(Enemy.expframe), Enemy.x * 32, (11 - Enemy.y + Npos) * 32 + Frameno
            Enemy.expframe = Enemy.expframe + 0.2
            If Enemy.expframe > 10 Then EnemyCol.Remove Enemy.Tag
        End If
    
    Next
    
    
    'Draw Shadow
    'Draw Car
    If Not dead Then
        DrawObject 11, CLng(x) + CLng(Z * 2), CLng(y)
        DrawObject 1, CLng(x), CLng(y), CLng(Z * 2)
    Else
        DrawOther CLng(expframe), CLng(x), CLng(y)
        expframe = expframe + 0.2
        If expframe > 10 Then expframe = 55
    End If
    
    DrawOther 17, 32, 64
    DrawOther 18, 64, 64
    If Gas > 48 Then
        DrawOther 11, 96, 64
        DrawOther 12, 128, 64, Gas * 4
        DrawOther 13, Gas * 4, 64
    Else
        DrawOther 14, 96, 64
        DrawOther 15, 128, 64, Gas * 4
        DrawOther 16, Gas * 4, 64
    End If
    
    a = Trim(Format(Str(Score), "00000"))
    For i = 0 To Len(a) - 1
        DrawOther 21 + Asc(Mid(a, i + 1, 1)) - 47, (i * 32) + 31, 32
    Next
    a = Trim(Format(Str(CLng(maxTime)), "000"))
    
    For i = 0 To Len(a) - 1
        DrawOther 21 + Asc(Mid(a, i + 1, 1)) - 47, (i * 32) + 380, 32
    Next
    
    If pause Then
        For i = 0 To 3
            DrawOther 33 + i, 220 + (i * 32), 200
        Next
    End If
    
    sRect.Top = 32
    sRect.Left = 32
    sRect.Right = 512
    sRect.Bottom = 384
    
    ddsPrimary.Blt wRect, ddsBuffer, sRect, DDBLT_WAIT
End Sub

Sub DrawObject(TileNum As Long, x As Long, y As Long, Optional Z As Long = 0, Optional rotate As Single = 0)
Dim dRect As RECT
Dim sRect As RECT

    dRect.Top = y - Z
    dRect.Left = x - Z
    dRect.Bottom = dRect.Top + 32 + Z
    dRect.Right = dRect.Left + 32 + Z
    
    sRect.Top = (TileNum \ 11) * 32
    sRect.Left = (TileNum Mod 11) * 32
    sRect.Right = sRect.Left + 32
    sRect.Bottom = sRect.Top + 32
 
    Dim dBltFX As DDBLTFX
    dBltFX.lRotationAngle = 90
    
    ddsBuffer.BltFx dRect, ddsObjects, sRect, DDBLT_DONOTWAIT + DDBLT_KEYSRC, dBltFX
    'ddsBuffer.Blt dRect, ddsObjects, sRect, DDBLT_DONOTWAIT + DDBLT_KEYSRC
End Sub


Sub DrawOther(TileNum As Long, x As Long, y As Long, Optional width As Variant)
Dim dRect As RECT
Dim sRect As RECT

    dRect.Top = y
    dRect.Left = x
    dRect.Bottom = dRect.Top + 32
    If IsMissing(width) Then
        dRect.Right = dRect.Left + 32
    Else
        dRect.Right = CLng(width)
    End If
    
    sRect.Top = (TileNum \ 11) * 32
    sRect.Left = (TileNum Mod 11) * 32
    sRect.Right = sRect.Left + 32
    sRect.Bottom = sRect.Top + 32
    
    ddsBuffer.Blt dRect, ddsOther, sRect, DDBLT_DONOTWAIT + DDBLT_KEYSRC
End Sub


Sub DrawRoad(TileNum As Long, x As Long, y As Long)
Dim dRect As RECT
Dim sRect As RECT

    dRect.Top = y
    dRect.Left = x
    dRect.Bottom = dRect.Top + 32
    dRect.Right = dRect.Left + 32
    
    sRect.Top = (TileNum \ 11) * 32
    sRect.Left = (TileNum Mod 11) * 32
    sRect.Right = sRect.Left + 32
    sRect.Bottom = sRect.Top + 32
    
    ddsBuffer.Blt dRect, ddsRoad, sRect, DDBLT_DONOTWAIT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bored = True
    Cancel = 1
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuInfo_Click()
    frmInfo.Show vbModal
End Sub

Private Sub mnureset_Click()
    dead = True
    newLevel
End Sub

Sub newLevel()
    If dead Then
        level = 0
        Score = 0
    End If
    
    'move to next level
    level = level + 1
    
    'load level data
    Open level & ".map" For Binary As 1
        Get #1, , RoadMap
        Get #1, , ObjectMap
        Get #1, , carx
    Close 1
    
    'clear enemies
    Set EnemyCol = Nothing
    
    nextlevel = False
    pause = False
    
    
    maxTime = 100
    
    my = 11 * 32
    mx = carx * 32
    mz = 0
    
    dead = False
    Gas = 100
    expframe = 0
    dx = 0
    dy = 0
    Npos = 0
    Frameno = 0
    deadtimer = 0
    keyStart = False
    
    Dim sRect As RECT
    Dim wRect As RECT
    
    Dim ry As Single
 
    If level = 1 Then
        While keyStart = False
            DoEvents
            
            sRect.Top = 32
            sRect.Left = 32
            sRect.Right = 512
            sRect.Bottom = 384
            
            wRect.Top = 0
            wRect.Left = 0
            wRect.Right = 512
            wRect.Bottom = 384
            
            ddsBuffer.Blt sRect, ddsBG, wRect, DDBLT_DONOTWAIT
            
            DX7.GetWindowRect picTarget.hWnd, wRect
            sRect.Top = 32
            sRect.Left = 32
            sRect.Right = 512
            sRect.Bottom = 384
            ddsPrimary.Blt wRect, ddsBuffer, sRect, DDBLT_WAIT
            
            Sleep 5
        If bored = True Then Exit Sub
        Wend
    End If
    
    'so we don't get that pause
    'when starting the game
    keyStart = False
End Sub

Private Sub picTarget_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyLeft
        keyLeft = True
    Case vbKeyRight
        keyRight = True
    Case vbKeyA
        keyAccel = True
    Case vbKeyZ
        keyBrake = True
    Case vbKeySpace
        keyFire = True
    Case vbKeyX
        keyJump = True
    Case vbKeyReturn
        keyStart = True
    End Select
End Sub

Private Sub picTarget_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyLeft
        keyLeft = False
    Case vbKeyRight
        keyRight = False
    Case vbKeyA
        keyAccel = False
    Case vbKeyZ
        keyBrake = False
    Case vbKeySpace
        keyFire = False
    Case vbKeyX
        keyJump = False
    Case vbKeyReturn
        keyStart = False
    End Select
End Sub




