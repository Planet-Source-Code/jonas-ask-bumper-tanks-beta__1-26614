Attribute VB_Name = "Weapons"
Public Type aThor
 Act As Boolean
 X As Integer
 Tag As Integer
 Owner As Byte
End Type
Public Thors(1 To 4) As aThor

Sub TankFire(A)
   With P(A)
    If .ReloadTime > 0 Then Exit Sub 'not reloaded, yet
    
    Dim i As Integer
    i = 1
    Do Until Shots(i).Act = False Or i = UBound(Shots)
        i = i + 1
    Loop
    If i = UBound(Shots) Then Exit Sub
    'It is ready...
    .FireSprite = 8
    .ReloadTime = 16
    PlaySound "fire" 'play sound
    If .Dire = 1 Then .Hspeed = .Hspeed + 1
    If .Dire = 2 Then .Hspeed = .Hspeed - 1
    
    Shots(i).Act = True
    Shots(i).Dis = 0
    Shots(i).Hspeed = IIf(.Dire = 1, -4, 4)
    Shots(i).VSpeed = -0.5
    Shots(i).X = IIf(.Dire = 1, .X + 1, .X + TankW - 1)
    Shots(i).Y = .Y + 1
    Shots(i).Owner = A 'set mother tank
    
   End With
End Sub
Public Sub MoveShots() 'This handels all weapon treatment
    'NORMAL SHOTS
    For A = 1 To UBound(Shots)
    With Shots(A)
        If .Act Then
            .Dis = .Dis + 4
            .X = .X + .Hspeed
            If .Dis > 60 Then .VSpeed = .VSpeed + 0.13
            .Y = .Y + .VSpeed
            ShotHitTerrain A
            ShotHitOpponent A
        End If
    End With
    Next A
    'THOR MISSILE
Dim FlagForKill As Boolean
    For A = 1 To UBound(Thors)
        If Thors(A).Act Then
            For B = 1 To 15
                Dim HitP As Integer
                Thors(A).Tag = Thors(A).Tag + 1
                For n = 1 To UBound(P)
                    If Thors(A).X >= P(n).X And Thors(A).X <= P(n).X + TankW Then
                    If Thors(A).Tag >= P(n).Y And Thors(A).Tag <= P(n).Y + TankH Then
                        HitP = n
                    End If
                    End If
                Next n
                If Main.PicBuffer(2).Point(Thors(A).X, Thors(A).Tag) <= 0 Or HitP > 0 Then
                    MakeExplo Thors(A).X - Int((Rnd * 15) + 3), Thors(A).Tag + Int(Rnd * 5) - 3
                    MakeExplo Thors(A).X, Thors(A).Tag + Int(Rnd * 5) - 3
                    MakeExplo Thors(A).X + Int((Rnd * 15) + 3), Thors(A).Tag + Int(Rnd * 5) - 3

                    If HitP > 0 Then AKilledB Thors(A).Owner, HitP
                        FlagForKill = True 'this allows it to destroy ALL tanks at target zone
                    Exit For
                End If
            Next B
            If FlagForKill Then
                Thors(A).Act = False
                Thors(A).Tag = 0
                Thors(A).X = 0
                Thors(A).Owner = 0
            End If
        End If
    Next A
End Sub
Public Sub FireSP(T)
    Select Case P(T).PUp.SPWeap
    Case 1 'Thor Missile
        ThorFire T
    
    
    End Select
    P(T).PUp.SPWeap = 0
End Sub
Public Sub ThorFire(A)
    Dim i As Integer
    i = 1
    Do Until Thors(i).Act = False Or i = UBound(Thors)
        i = i + 1
    Loop
    
    PlaySound "thor"
    'get a random opponent to hit
newO:
    O = Int((Rnd * UBound(P)) + 1)
    If O = A Then GoTo newO
    
    Thors(i).Act = True
    Thors(i).X = P(O).X + (TankW \ 2)
    Thors(i).Tag = 0
    Thors(i).Owner = A
End Sub
