Attribute VB_Name = "Publics2"
Public MsgString As String
Public MsgTag As Integer

Public Function Checkforstop(n) As Boolean
Dim PixU() As Long, PixO() As Long
Dim Suported As Byte
    ReDim PixU(1 To 6)
    ReDim PixO(1 To 3)
    With P(n)
        Checkforstop = False
        
        For B = 1 To 6
            PixU(B) = Main.PicBuffer(2).Point(.X + B + 4, .Y + TankH)
            If PixU(B) = 0 Then Suported = Suported + 1
        Next B
        PixO(1) = Main.PicBuffer(2).Point(.X + 2, .Y)
        PixO(2) = Main.PicBuffer(2).Point(.X + 9, .Y)
        PixO(3) = Main.PicBuffer(2).Point(.X + TankW - 2, .Y)
        
        If Suported > 0 Then
            Checkforstop = True
        End If
        If PixO(1) = 0 Or PixO(2) = 0 Or PixO(3) = 0 Then
            Checkforstop = True
            .Y = .Y + 2
            .VSpeed = -.VSpeed
        End If
    End With
End Function
Public Function CheckforstopX(n) As Boolean
Dim PixL() As Long, PixR() As Long
    ReDim PixL(1 To 7)
    ReDim PixR(1 To 7)
    With P(n)
        CheckforstopX = False
        
        For B = 1 To UBound(PixR)
            PixL(B) = Main.PicBuffer(2).Point(.X + 3, .Y + B)
            PixR(B) = Main.PicBuffer(2).Point(.X + TankW - 3, .Y + B)
            If PixL(B) = 0 Then OpenR = OpenR + 1
            If PixR(B) = 0 Then OpenL = OpenL + 1
        Next B
        
        If OpenR > 0 Then
            CheckforstopX = True
            .Hspeed = -.Hspeed
            .X = .X + 2
            Exit Function
        End If
        If OpenL > 0 Then
            CheckforstopX = True
            .Hspeed = -.Hspeed
            .X = .X - 2
            Exit Function
        End If
        If .Dire = 2 Then
            If Main.PicBuffer(2).Point(.X + TankW - 3, .Y + 9) = 0 Then
                .Y = .Y - 1
                Exit Function
            End If
        Else
            If Main.PicBuffer(2).Point(.X + 3, .Y + 9) = 0 Then
                .Y = .Y - 1
                Exit Function
            End If
        End If
    End With
End Function

Public Sub SpawnTank(A)
    P(A).X = Int((Rnd * (BoardW - 160)) + 80)
    P(A).Y = Int((Rnd * 30) + 30)
    P(A).Dire = Int(Rnd * 2) + 1
    P(A).Life = 2
    P(A).VSpeed = Rnd * 4 - 6
    P(A).Hspeed = Rnd * 4 - 2
    P(A).PUp.JetPack = 0
    P(A).PUp.SPWeap = 0
End Sub

Public Sub DoCrates()
    If Rnd * 10000 > 9980 Then MakeCrate
        
    For A = 1 To UBound(Crates)
        If Crates(A).Act Then
            If Crates(A).Timeleft > 0 Then
                Crates(A).Timeleft = Crates(A).Timeleft - 1
            Else
                Crates(A).Act = False
                Crates(A).Cont = 0
                Crates(A).X = 0
                Crates(A).Y = 0
            End If
            'Check for player pick up
            For B = 1 To UBound(P)
                If P(B).X + TankW >= Crates(A).X And P(B).X <= Crates(A).X + 15 Then
                If P(B).Y + TankH >= Crates(A).Y And P(B).Y <= Crates(A).Y + 15 Then
                    PickUpcrate Crates(A).Cont, B
                    Crates(A).Act = False
                    Crates(A).Cont = 0
                    Crates(A).X = 0
                    Crates(A).Y = 0
                    Exit For
                End If
                End If
            Next B
        End If
    Next A
End Sub

Sub MakeCrate()
    Dim i As Integer
    i = 1
    Do Until Crates(i).Act = False Or i = UBound(Crates)
        i = i + 1
    Loop
    Crates(i).Act = True 'Activate this crate
    Crates(i).Timeleft = 300
GetOther:
    X = Int((Rnd * (BoardW - 80) + 40)) 'Find some OK coords
    Y = Int((Rnd * (BoardH - 60) + 40))
    If Not IsCrateOK(X, Y) Then GoTo GetOther
    
    Crates(i).X = X 'assign the found coords
    Crates(i).Y = Y
    
    Crates(i).Cont = Int((Rnd * 4) + 1) 'Create contense
End Sub
Function IsCrateOK(X, Y) As Boolean
Dim PixU() As Long, PixO() As Long
Dim Suported As Byte
Dim Over As Byte
    IsCrateOK = False
    ReDim PixU(1 To 6)
    ReDim PixO(1 To 6)
        For B = 1 To 6
            PixU(B) = Main.PicBuffer(2).Point(X + B + 4, Y + 15) '15 is the height of a crate
            PixO(B) = Main.PicBuffer(2).Point(X + B + 4, Y + 14)
            If PixU(B) = 0 Then Suported = Suported + 1
            If PixO(B) = 0 Then Over = Over + 1
        Next B
        If Suported > 0 And Over = 0 Then IsCrateOK = True
End Function
Sub PickUpcrate(C, A) 'C is Contents, a is player
    Select Case C
    Case 1 'Jetpack
        P(A).PUp.JetPack = P(A).PUp.JetPack + 70
        CreateMessage "Jetpack", A
    Case 2 'Health kit
        P(A).Life = P(A).Life + 2
        If P(A).Life > 6 Then P(A).Life = 6
        CreateMessage "Health kit", A
    Case 3 'Super Speed
        P(A).PUp.SuperSpeed = 600
        CreateMessage "Super Fuel", A
    Case 4 'THOR Missile
        P(A).PUp.SPWeap = 1
        CreateMessage "THOR Missile", A
    End Select
End Sub
Public Sub CreateMessage(Txt, Col)
Dim C As Long
    Select Case Col
    Case 1: C = RGB(170, 170, 255)
    Case 2: C = vbGreen
    Case 0: C = vbWhite
    End Select
    MsgTag = 40
    MsgString = Txt
    Main.lblInfo.ForeColor = C
End Sub
Public Sub DoMessage()
    If MsgTag > 0 Then
        MsgTag = MsgTag - 1
    Else
        MsgString = ""
    End If
    Main.lblInfo.Caption = MsgString
End Sub

