Attribute VB_Name = "Publics"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCCOPY = &HCC0020

Public Type Powerups
 JetPack As Integer
 SuperSpeed As Integer
 SPWeap As Byte
End Type

Public Type Player
 X As Currency
 Y As Currency
 VSpeed As Currency
 Hspeed As Currency
 Dire As Byte
 PUp As Powerups
 FireSprite As Integer
 Suport As Boolean
 Ammo As Byte
 ReloadTime As Byte
 Points As Integer
 Life As Byte
End Type

Public Type Shot
 X As Currency
 Y As Currency
 VSpeed As Currency
 Hspeed As Currency
 Dis As Currency
 Act As Boolean
 Owner As Integer
End Type

Public Type aCrate
 Act As Boolean
 X As Integer
 Y As Integer
 Cont As Integer
 Timeleft As Integer
End Type

Public Type misc
 X As Currency
 Y As Currency
 Act As Boolean
 Tag As Integer
 Speed As Currency
End Type

Public Explo(1 To 10) As misc
Public Clouds(1 To 10) As misc
Public Crates(1 To 10) As aCrate
Public Shots(1 To 20) As Shot
Public P(1 To 2) As Player
Public Const TankW = 18, TankH = 10
Public Const MaxH = 3, MaxV = 8
Public BoardW As Integer, BoardH As Integer
Public TheKing As Integer
Public MainPause As Boolean

Public Sub LoadPictures()
    Main.PicMain.Width = Main.PicBuffer(1).Width
    Main.PicMain.Height = Main.PicBuffer(1).Height
    For A = 0 To Main.PicBuffer.UBound
        Main.PicBuffer(A).Width = Main.PicMain.Width
        Main.PicBuffer(A).Height = Main.PicMain.Height
    Next A
    BoardH = Main.PicMain.ScaleHeight
    BoardW = Main.PicMain.ScaleWidth
End Sub

Public Sub DoGraphics()
    With Main
    .PicMain.Cls
    .PicBuffer(0).Cls
    .PicBuffer(1).Cls
    'The background
    If TheKing > 0 Then
        BitBlt .PicBuffer(0).hDC, TheKing - 148, 2, 148, 186, .PicKingM.hDC, 0, 0, SRCAND
        BitBlt .PicBuffer(0).hDC, TheKing - 148, 2, 148, 186, .PicKing.hDC, 0, 0, SRCPAINT
    End If
    For A = 1 To UBound(Clouds)
        If Clouds(A).Act Then
            BitBlt .PicBuffer(0).hDC, Clouds(A).X, Clouds(A).Y, .PicCloud(Clouds(A).Tag).ScaleWidth, .PicCloud(Clouds(A).Tag).ScaleHeight, .PicCloudM(Clouds(A).Tag).hDC, 0, 0, SRCAND
            BitBlt .PicBuffer(0).hDC, Clouds(A).X, Clouds(A).Y, .PicCloud(Clouds(A).Tag).ScaleWidth, .PicCloud(Clouds(A).Tag).ScaleHeight, .PicCloud(Clouds(A).Tag).hDC, 0, 0, SRCPAINT
        End If
    Next A
    BitBlt .PicBuffer(0).hDC, 0, 0, BoardW, BoardH, .PicBackM.hDC, 0, 0, SRCAND
    BitBlt .PicBuffer(0).hDC, 0, 0, BoardW, BoardH, .PicBack.hDC, 0, 0, SRCPAINT
    
    BitBlt .PicBuffer(0).hDC, 0, 0, BoardW, BoardH, .PicBuffer(2).hDC, 0, 0, SRCAND
    BitBlt .PicBuffer(0).hDC, 0, 0, BoardW, BoardH, .PicBuffer(1).hDC, 0, 0, SRCPAINT
    'The crates
    For A = 1 To UBound(Crates)
        If Crates(A).Act Then
            BitBlt .PicBuffer(0).hDC, Crates(A).X, Crates(A).Y, 15, 15, .PicCrate.hDC, 0, 0, SRCCOPY
        End If
    Next A
    'Paint the players
    For A = 1 To UBound(P)
        T = IIf(P(A).Dire = 1, 0, 1)
        BitBlt .PicBuffer(0).hDC, P(A).X, P(A).Y, TankW, TankH, .PicPlayerM(P(A).Dire - 1).hDC, 0, 0, SRCAND
        BitBlt .PicBuffer(0).hDC, P(A).X, P(A).Y, TankW, TankH, .PicPlayer(T).hDC, 0, 0, SRCPAINT
        If P(A).FireSprite > 0 Then
            If P(A).Dire = 1 Then
                BitBlt .PicBuffer(0).hDC, P(A).X - 12, P(A).Y, 12, 4, .PicMuzM(Int(P(A).FireSprite / 4)).hDC, 0, 0, SRCAND
                BitBlt .PicBuffer(0).hDC, P(A).X - 12, P(A).Y, 12, 4, .PicMuz(Int(P(A).FireSprite / 4)).hDC, 0, 0, SRCPAINT
            Else
                BitBlt .PicBuffer(0).hDC, P(A).X + 18, P(A).Y, 12, 4, .PicMuzM(Int(P(A).FireSprite / 4) + 2).hDC, 0, 0, SRCAND
                BitBlt .PicBuffer(0).hDC, P(A).X + 18, P(A).Y, 12, 4, .PicMuz(Int(P(A).FireSprite / 4) + 2).hDC, 0, 0, SRCPAINT
            End If
        End If
    Next A
    'The Shells
    For A = 1 To UBound(Shots)
        If Shots(A).Act Then
            BitBlt .PicBuffer(0).hDC, Shots(A).X, Shots(A).Y, 4, 4, .PicShellM.hDC, 0, 0, SRCAND
            BitBlt .PicBuffer(0).hDC, Shots(A).X, Shots(A).Y, 4, 4, .PicShell.hDC, 0, 0, SRCPAINT
        End If
    Next A
    'The THOR
    For A = 1 To UBound(Thors)
        If Thors(A).Act Then
            .PicBuffer(0).Line (Thors(A).X - 3, 0)-Step(6, Thors(A).Tag), RGB(255, 63, 19), BF
            .PicBuffer(0).Line (Thors(A).X - 1, 0)-Step(2, Thors(A).Tag + 5), vbRed, BF
        End If
    Next A
    'The explosions
    For A = 1 To UBound(Explo)
        If Explo(A).Act Then
            BitBlt .PicBuffer(0).hDC, Explo(A).X, Explo(A).Y, 20, 18, .PicExpM(Int(Explo(A).Tag / 3)).hDC, 0, 0, SRCAND
            BitBlt .PicBuffer(0).hDC, Explo(A).X, Explo(A).Y, 20, 18, .PicExp(Int(Explo(A).Tag / 3)).hDC, 0, 0, SRCPAINT
        End If
    Next A

    
    BitBlt .PicMain.hDC, 0, 0, BoardW, BoardH, .PicBuffer(0).hDC, 0, 0, SRCCOPY
    End With
End Sub
Public Sub DoPhysics()
Dim PixU() As Long
Dim Suported As Byte
    ReDim PixU(1 To 6)
    For A = 1 To UBound(P)
    With P(A)
        Suported = 0
        
        For B = 1 To 6
            PixU(B) = Main.PicBuffer(2).Point(.X + B + 4, .Y + TankH)
            If PixU(B) = 0 Then Suported = Suported + 1
        Next B
        If Suported = 0 Then 'free fall
            .VSpeed = .VSpeed + 0.4
            .Hspeed = .Hspeed / 1.05
            .Suport = False
        Else 'hit ground
            .VSpeed = 0
            .Suport = True
            .Hspeed = .Hspeed / 1.1
        End If
        
        
        If .VSpeed > MaxV Then .VSpeed = MaxV
        If .VSpeed < -MaxV Then .VSpeed = -MaxV
        If P(A).PUp.SuperSpeed > 0 Then
            If .Hspeed > MaxH * 2 Then .Hspeed = MaxH * 2
            If .Hspeed < -(MaxH * 2) Then .Hspeed = -(MaxH * 2)
        Else
            If .Hspeed > MaxH Then .Hspeed = MaxH
            If .Hspeed < -MaxH Then .Hspeed = -MaxH
        End If
        MoveTank A
        
    End With
    Next A
End Sub
Sub MoveTank(B)
Dim PixU() As Long
Dim Suported As Byte
    ReDim PixU(1 To 6)
    If P(B).Hspeed > 0 Then
        For A = 0.1 To P(B).Hspeed Step 0.1
            If CheckforstopX(B) Then Exit For
            P(B).X = P(B).X + 0.1
        Next A
    Else
        For A = -0.1 To P(B).Hspeed Step -0.1
            If CheckforstopX(B) Then Exit For
            P(B).X = P(B).X - 0.1
        Next A
    End If
    If P(B).VSpeed > 0 Then
        For A = 0.1 To P(B).VSpeed Step 0.1
            If Checkforstop(B) Then Exit For
            P(B).Y = P(B).Y + 0.1
        Next A
    Else
        For A = -0.1 To P(B).VSpeed Step -0.1
            If Checkforstop(B) Then Exit For
            P(B).Y = P(B).Y - 0.1
        Next A
    End If
    'Check if off screen
    If P(B).X < -10 Or P(B).X > BoardW + 1 Or P(B).Y < -60 Or P(B).Y > BoardH + 10 Then
        P(B).Points = P(B).Points - 10
        SpawnTank B
    End If
    'Power ups
    If P(B).PUp.SuperSpeed > 0 Then
        P(B).PUp.SuperSpeed = P(B).PUp.SuperSpeed - 1
    End If
End Sub
Public Sub TankThink()
    For A = 1 To UBound(P)
    With P(A)
        If .FireSprite > 0 Then
            .FireSprite = .FireSprite - 1
        End If
        
        If .ReloadTime > 0 Then
            .ReloadTime = .ReloadTime - 1
        End If
        
    End With
    Next A
    
    'The king sneaks in
    If TheKing > 0 Then
        TheKing = TheKing - 4
    End If
End Sub
Sub TankJump(A)
    P(A).VSpeed = -5
    P(A).Y = P(A).Y - 1
End Sub
Sub ShotHitOpponent(s)
    For A = 1 To UBound(P)
        If Shots(s).X + 2 >= P(A).X And Shots(s).X - 2 <= P(A).X + TankW Then
        If Shots(s).Y >= P(A).Y And Shots(s).Y <= P(A).Y + TankH Then
        If A <> Shots(s).Owner Then 'don't hit owner
            O = Shots(s).Owner
            MakeExplo Shots(s).X, Shots(s).Y
            If Shots(s).X < P(A).X + (TankW / 2) Then
                P(A).Hspeed = P(A).Hspeed - 2
            Else
                P(A).Hspeed = P(A).Hspeed + 2
            End If
            If P(A).Life > 0 Then
                P(A).Life = P(A).Life - 1
            Else
                AKilledB O, A
            End If
            P(A).VSpeed = P(A).VSpeed - 2
            Shots(s).Act = False
            Shots(s).Dis = 0
            Shots(s).Hspeed = 0
            Shots(s).Owner = 0
            Shots(s).VSpeed = 0
            Shots(s).X = 0
            Shots(s).Y = 0
        End If
        End If
        End If
    Next A
End Sub
Sub ShotHitTerrain(s)
Dim Pix(1 To 4)
    Pix(1) = Main.PicBuffer(2).Point(Shots(s).X + 4, Shots(s).Y)
    Pix(2) = Main.PicBuffer(2).Point(Shots(s).X, Shots(s).Y + 4)
    Pix(3) = Main.PicBuffer(2).Point(Shots(s).X + 4, Shots(s).Y + 4)
    Pix(4) = Main.PicBuffer(2).Point(Shots(s).X, Shots(s).Y)
    
    If Pix(1) = 0 Or Pix(2) = 0 Or Pix(3) = 0 Or Pix(4) = 0 Then 'hit something
        MakeExplo Shots(s).X, Shots(s).Y
        Shots(s).Act = False
        Shots(s).Dis = 0
        Shots(s).Hspeed = 0
        Shots(s).Owner = 0
        Shots(s).VSpeed = 0
        Shots(s).X = 0
        Shots(s).Y = 0
    End If
    If Shots(s).X <= 0 Or Shots(s).X >= BoardW Or Shots(s).Y <= 0 Or Shots(s).Y >= BoardH Then
        Shots(s).Act = False
        Shots(s).Dis = 0
        Shots(s).Hspeed = 0
        Shots(s).Owner = 0
        Shots(s).VSpeed = 0
        Shots(s).X = 0
        Shots(s).Y = 0
    End If
End Sub
Sub MakeExplo(tX, ty)
Dim X As Integer, Y As Integer
    X = tX: Y = ty
    'play a sound
    PlaySound "explo"
    X = X - 10
    Y = Y - 9
    A = 1
    Do Until Not Explo(A).Act Or A = UBound(Explo)
        A = A + 1
    Loop
    With Explo(A)
        .X = X
        .Y = Y
        .Tag = 0
        .Act = True
    End With
End Sub
Public Sub DoExplo()
    For A = 1 To UBound(Explo)
        If Explo(A).Act Then
            If Explo(A).Tag < 11 Then
                Explo(A).Tag = Explo(A).Tag + 1
            Else
                Explo(A).Act = False
                Explo(A).X = 0
                Explo(A).Y = 0
                Explo(A).Tag = 0
            End If
        End If
    Next A
End Sub
Public Sub DoKeys()
    Rateair = 0.1
    rateground = 0.3
    sp1 = IIf(P(1).PUp.SuperSpeed > 0, True, False)
    sp2 = IIf(P(2).PUp.SuperSpeed > 0, True, False)
    'Player 1
    If GetAsyncKeyState(vbKeyLeft) Then
        If P(1).Suport Then
            P(1).Hspeed = P(1).Hspeed - IIf(sp1, rateground * 2, rateground)
        Else
            P(1).Hspeed = P(1).Hspeed - Rateair
        End If
        P(1).Dire = 1
    End If
    If GetAsyncKeyState(vbKeyRight) Then
        If P(1).Suport Then
            P(1).Hspeed = P(1).Hspeed + IIf(sp1, rateground * 2, rateground)
        Else
            P(1).Hspeed = P(1).Hspeed + Rateair
        End If
        P(1).Dire = 2
    End If
    If GetAsyncKeyState(vbKeyUp) And (P(1).Suport Or P(1).PUp.JetPack > 0) Then
        If P(1).PUp.JetPack > 0 Then P(1).PUp.JetPack = P(1).PUp.JetPack - 1
        TankJump 1
    End If
    If GetAsyncKeyState(vbKeyNumpad0) Then
        TankFire 1
    End If
    If GetAsyncKeyState(vbKeyNumpad1) Then
        FireSP 1
    End If
    'Player 2
    If GetAsyncKeyState(vbKeyA) Then
        If P(2).Suport Then
            P(2).Hspeed = P(2).Hspeed - IIf(sp2, rateground * 2, rateground)
        Else
            P(2).Hspeed = P(2).Hspeed - Rateair
        End If
        P(2).Dire = 1
    End If
    If GetAsyncKeyState(vbKeyD) Then
        If P(2).Suport Then
            P(2).Hspeed = P(2).Hspeed + IIf(sp2, rateground * 2, rateground)
        Else
            P(2).Hspeed = P(2).Hspeed + Rateair
        End If
        P(2).Dire = 2
    End If
    If GetAsyncKeyState(vbKeyW) And (P(2).Suport Or P(2).PUp.JetPack > 0) Then
        If P(2).PUp.JetPack > 0 Then P(2).PUp.JetPack = P(2).PUp.JetPack - 1
        TankJump 2
    End If
    If GetAsyncKeyState(vbKeyQ) Then
        TankFire 2
    End If
    If GetAsyncKeyState(vbKeyE) Then
        FireSP 2
    End If

End Sub
Public Function PlaySound(File As String)
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
    'King Override
    If TheKing > 0 Then Exit Function
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Svar = sndPlaySound(App.Path & "\" & File & ".wav", wFlags%) 'Send the sound to the big world
End Function

Public Sub DoClouds()
    If Rnd > 0.99 Then
        A = 1
        Do Until Not Clouds(A).Act Or A = UBound(Clouds)
            A = A + 1
        Loop
        If A <= UBound(Clouds) Then
            Clouds(A).Act = True
            Clouds(A).Tag = Int(Rnd * 5)
            Clouds(A).Speed = IIf(Rnd > 0.5, Int((Rnd * 4) + 2), -Int((Rnd * 4) + 2))
            Clouds(A).X = IIf(Clouds(A).Speed < 0, BoardW, -Main.PicCloud(Clouds(A).Tag).ScaleWidth)
            Clouds(A).Y = Int(Rnd * 180) - 10
        End If
    End If
    For A = 1 To UBound(Clouds)
        If Clouds(A).Act Then
            Clouds(A).X = Clouds(A).X + Clouds(A).Speed
            If Clouds(A).X <= -3 - Main.PicCloud(Clouds(A).Tag).ScaleWidth Or Clouds(A).X >= BoardW + 3 Then
                Clouds(A).Act = False
                Clouds(A).Speed = 0
                Clouds(A).Tag = 0
                Clouds(A).X = 0
                Clouds(A).Y = 0
            End If
        End If
    Next A
End Sub

Public Sub AKilledB(A, B)
    CreateMessage "Player " & B & " destroyed", B
    If A <> B Then P(A).Points = P(A).Points + 1 'Do not give points if we hit ourself
    SpawnTank B
End Sub
