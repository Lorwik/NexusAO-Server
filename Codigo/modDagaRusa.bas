Attribute VB_Name = "modDagaRusa"
Option Explicit

Private Const TIEMPO_CANCELACION As Integer = 180

Public Const NPC_DAGA_RUSA As Integer = 1050
Public INDEX_NPC_DAGA_RUSA_ONLINE As Integer

Private Type tUsuario
    ID As Integer
    Posicion As WorldPos
End Type

Private Type tDagaRusa
    Activo As Boolean
    Usuarios() As tUsuario
    Conteo As Integer
    Cupos As Byte
    CoordenadasEspera As WorldPos
    CoordenadasArena As WorldPos
    CoordenadasNPC As WorldPos
    Premio As Long
    Inscripcion As Long
    Total As Byte
    Restantes As Byte
    AtacoUser As Integer
    Atacar As Integer
    PuedeAtacar As Boolean
    ActivarEvento As Boolean
    Volver As Boolean
End Type

Private DagaRusa As tDagaRusa

Public Sub InitDagaRusa()

    With DagaRusa.CoordenadasArena
        .Map = CInt(Leer.GetValue("EVENTO", "Mapa_Espera"))
        .X = CByte(Leer.GetValue("EVENTO", "X_Espera"))
        .Y = CByte(Leer.GetValue("EVENTO", "Y_Espera"))
    End With

    With DagaRusa.CoordenadasEspera
        .Map = CInt(Leer.GetValue("EVENTO", "Mapa_Arena"))
        .X = CByte(Leer.GetValue("EVENTO", "X_Arena"))
        .Y = CByte(Leer.GetValue("EVENTO", "Y_Arena"))
    End With

End Sub

Public Sub Armar_DagaRusa(ByVal ID As Integer, ByVal Cupos As Byte, ByVal Premio As Long, ByVal Inscripcion As Long)

    With DagaRusa
        If .Activo = True Then
            Call WriteConsoleMsg(ID, "Daga Rusa> El evento ya está en curso.", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If

        If Cupos > 16 Then Cupos = 16
        If Cupos < 2 Then Cupos = 2
        If Premio <= 0 Then Premio = 1
       
        .Cupos = Cupos
        .Inscripcion = Inscripcion
        .Premio = Premio
        .Total = .Cupos
        .Restantes = .Total
        .Activo = True
        .Conteo = TIEMPO_CANCELACION
        ReDim .Usuarios(1 To .Cupos) As tUsuario
       
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Daga Rusa> " & .Cupos & " Cupos, Incripción" & IIf(.Inscripcion > 0, " de: " & .Inscripcion & " Monedas de oro, ", " Gratis, ") & IIf(.Premio > 0, "Premio de: " & .Premio & " Monedas de oro.", " No hay premio.") & " Manden /DAGARUSA si desean participar.", FontTypeNames.FONTTYPE_GUILD))
    End With

End Sub

Public Sub Entrar_DagaRusa(ByVal Userindex As Integer)

    Dim ID_DagaRusa As Byte
    Dim LoopC As Long

    With DagaRusa
        If Puede_Entrar(Userindex) = False Then _
           Exit Sub

        Call WriteConsoleMsg(Userindex, "Has ingresado al evento" & IIf(.Inscripcion > 0, ", se te han descontado " & .Inscripcion & " monedas de oro.", vbNullString) & ". Espera a que el cupo se complete. ¡Suerte en el campo de batalla!", FontTypeNames.FONTTYPE_GUILD)

        UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - .Inscripcion
        ID_DagaRusa = DagaRusa_ID
        UserList(Userindex).flags.EnDagaRusa = ID_DagaRusa

        .Cupos = .Cupos - 1
        .Usuarios(ID_DagaRusa).ID = ID
        .Usuarios(ID_DagaRusa).Posicion = UserList(Userindex).Pos

        With DagaRusa.CoordenadasEspera
            WarpUserChar Userindex, .Map, .X, .Y, False
        End With

        WritePauseToggle Userindex
        WriteUpdateGold Userindex

        If .Cupos = 0 Then
            For LoopC = 1 To .Total
                WarpUserChar .Usuarios(LoopC).ID, .CoordenadasArena.Map, .CoordenadasArena.X + LoopC, .CoordenadasArena.Y, True
            Next LoopC
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Daga Rusa> El cupo ha sido completado!", FontTypeNames.FONTTYPE_GUILD))
            .ActivarEvento = True
            .Conteo = 10
            .CoordenadasNPC = UserList(.Usuarios(1).ID).Pos
            .CoordenadasNPC.Y = .CoordenadasNPC.Y - 1
            SpawnNpc NPC_DAGA_RUSA, .CoordenadasNPC, False, False
        End If
    End With

End Sub

Private Function DagaRusa_ID() As Byte

    Dim LoopC As Long

    With DagaRusa
        For LoopC = 1 To .Total
            If .Usuarios(LoopC).ID = 0 Then
                DagaRusa_ID = LoopC
                Exit Function
            End If
        Next LoopC
    End With

End Function

Private Function Puede_Entrar(ByVal ID As Integer) As Boolean

    Puede_Entrar = False

    If UserList(ID).flags.Muerto > 0 Then
        Call WriteConsoleMsg(ID, "Estás muerto.", FontTypeNames.FONTTYPE_GUILD)
        Exit Function
    End If

    'If UserList(ID).flags.EnJDH > 0 Then
    '    Call WriteConsoleMsg(ID, "Estás en los Juegos del Hambre.", FontTypeNames.FONTTYPE_GUILD)
    '    Exit Function
    'End If

    'If UserList(ID).flags.EnPlantes > 0 Then
    '    Call WriteConsoleMsg(ID, "Ya estás en Plantes Automáticos.", FontTypeNames.FONTTYPE_GUILD)
    '    Exit Function
    'End If

    If UserList(ID).flags.EnDagaRusa > 0 Then
        Call WriteConsoleMsg(ID, "Ya estás en el en Daga Rusa.", FontTypeNames.FONTTYPE_GUILD)
        Exit Function
    End If

    If DagaRusa.Activo = False Then
        Call WriteConsoleMsg(ID, "El evento no está en curso.", FontTypeNames.FONTTYPE_GUILD)
        Exit Function
    End If

    If DagaRusa.Cupos = 0 Then
        Call WriteConsoleMsg(ID, "El evento ya no tiene cupos disponibles.", FontTypeNames.FONTTYPE_GUILD)
        Exit Function
    End If

    If UserList(ID).Stats.Gld < DagaRusa.Inscripcion Then
        Call WriteConsoleMsg(ID, "No tienes el oro suficiente.", FontTypeNames.FONTTYPE_GUILD)
        Exit Function
    End If

    If Not UserList(ID).Pos.Map = 1 Then
        Call WriteConsoleMsg(ID, "Tienes que estar en Ullathorpe para poder ingresar al evento", FontTypeNames.FONTTYPE_GUILD)
        Exit Function
    End If
   
    If Tiene_Objeto(ID) = False Then
        Call WriteConsoleMsg(ID, "No tienes que tener ningún objeto en tu inventario para ingresar al evento.", FontTypeNames.FONTTYPE_GUILD)
        'Exit Function
    End If

    Puede_Entrar = True

End Function

Public Sub Contar_DagaRusa()

    Dim LoopC As Long
    Dim LoopX As Long
    Dim ID_DagaRusa As Byte

    With DagaRusa
        If .Conteo = 0 Then
            .Conteo = -1
            If .Activo = True Then
                If .ActivarEvento = True Then
                    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Daga Rusa> Ya!", FontTypeNames.FONTTYPE_FIGHT)
                    .PuedeAtacar = True
                Else
                    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Daga Rusa> Evento cancelado por falta de participantes, se ha devuelto el oro por la inscripción.", FontTypeNames.FONTTYPE_GUILD)
                    Cancelar_DagaRusa
                End If
            End If
        End If
    '        .Conteo = 3
        If .Conteo > 0 Then
            If .ActivarEvento = True Then _
               SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Daga Rusa> " & .Conteo, FontTypeNames.FONTTYPE_GUILD)
            .Conteo = .Conteo - 1
        End If
    End With
End Sub

Public Sub IA_NPC_DAGARUSA(ByVal NpcIndex As Integer)

    Dim Y As Long
    Dim X As Long
    Dim UI As Integer
    Dim tHeading As Byte

    With Npclist(NpcIndex)
        If DagaRusa.PuedeAtacar = True Then
            If DagaRusa.Atacar > 0 Then
                NpcAtacaUser NpcIndex, DagaRusa.Atacar
                DagaRusa.AtacoUser = DagaRusa.Atacar
                DagaRusa.Atacar = 0
            End If

            For Y = .Pos.Y To .Pos.Y + RANGO_VISION_Y
                For X = .Pos.X To .Pos.X + RANGO_VISION_Y
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        UI = MapData(.Pos.Map, X, Y).Userindex
                        If UI > 0 Then
                            If UI <> DagaRusa.AtacoUser Then
                                If DagaRusa.Volver = False Then
                                    If Distancia(.Pos, UserList(UI).Pos) <= 1 Then
                                        If DagaRusa.Atacar = 0 Then
                                            DagaRusa.Atacar = UI
                                            .Char.Heading = SOUTH
                                            ChangeNPCChar NpcIndex, .Char.body, .Char.Head, .Char.Heading
                                            Exit Sub
                                        End If
                                    End If
                                    If UserList(UI).flags.EnDagaRusa = Total Then DagaRusa.Volver = True
                                    tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(UI).Pos)
                                Else
                                    tHeading = FindDirection(Npclist(NpcIndex).Pos, DagaRusa.CoordenadasNPC)
                                End If
                                MoveNPCChar NpcIndex, tHeading
                                Exit Sub
                            End If
                        End If
                    End If
                Next X
            Next Y
        End If
    End With
End Sub

Private Function ID_Usuario() As Byte

    Dim LoopC As Long

    For LoopC = 1 To DagaRusa.Total
        If DagaRusa.Usuarios(LoopC).ID > 0 Then
            ID_Usuario = LoopC
            Exit For
        End If
    Next LoopC

End Function

Public Sub Apuñalado_DagaRusa(ByVal ID As Integer)

    Dim ID_DagaRusa As Byte

    ID_DagaRusa = UserList(ID).flags.EnDagaRusa
    UserList(ID).flags.EnDagaRusa = 0

    With DagaRusa
        .Restantes = .Restantes - 1
        If .Restantes > 1 Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Daga Rusa> Quedan " & .Restantes & " participantes.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(ID, "Daga Rusa> ¡Has perdido, has sido descalificado. ¡Suerte para la próxima!", FontTypeNames.FONTTYPE_GUILD)
        WarpUserChar ID, .Usuarios(ID_DagaRusa).Posicion.Map, .Usuarios(ID_DagaRusa).Posicion.X, .Usuarios(ID_DagaRusa).Posicion.Y, False
        .Usuarios(ID_DagaRusa).ID = 0
        If .Restantes = 1 Then
            Call QuitarNPC(INDEX_NPC_DAGA_RUSA_ONLINE)
            Call Finalizar
        End If
    End With

End Sub

Private Sub Finalizar()

    Dim LoopC As Long
    Dim Dame_ID As Byte
    Dim ID As Integer

    With DagaRusa
        Dame_ID = ID_Usuario
        ID = .Usuarios(Dame_ID).ID
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Daga Rusa> Ganador del evento: " & UserList(ID).name & " se lleva una cantidad de " & .Premio & " monedas de oro, felicitaciones!", FontTypeNames.FONTTYPE_GUILD)
        UserList(ID).Stats.Gld = UserList(ID).Stats.Gld + .Premio

        WriteUpdateGold ID
        UserList(ID).flags.EnDagaRusa = 0
        .Premio = 0
        WarpUserChar ID, .Usuarios(Dame_ID).Posicion.Map, .Usuarios(Dame_ID).Posicion.X, .Usuarios(Dame_ID).Posicion.Y, False
        Call Limpiar
    End With
End Sub

Public Sub Cancelar_DagaRusa()

    Dim LoopC As Long

    With DagaRusa

        If .Activo = False Then Exit Sub

        For LoopC = 1 To .Total
            If .Usuarios(LoopC).ID > 0 Then
                WarpUserChar .Usuarios(LoopC).ID, .Usuarios(LoopC).Posicion.Map, .Usuarios(LoopC).Posicion.X, .Usuarios(LoopC).Posicion.Y, False
                UserList(.Usuarios(LoopC).ID).flags.EnDagaRusa = 0
                UserList(.Usuarios(LoopC).ID).Stats.Gld = UserList(.Usuarios(LoopC).ID).Stats.Gld + .Inscripcion

                WriteConsoleMsg .Usuarios(LoopC).ID, "El evento ha sido cancelado, se te ha devuelto el costo de la inscripción.", FontTypeNames.FONTTYPE_GUILD
                WriteUpdateGold .Usuarios(LoopC).ID
            End If
        Next LoopC
    End With

    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Daga Rusa> Evento fue cancelado por un Game Master.", FontTypeNames.FONTTYPE_GUILD)
    Limpiar

End Sub

Public Sub Desconexion_DagaRusa(ByVal ID As Integer)

    If UserList(ID).flags.EnDagaRusa = 0 Then Exit Sub

    With DagaRusa
        WarpUserChar ID, .Usuarios(UserList(ID).flags.EnDagaRusa).Posicion.Map, .Usuarios(UserList(ID).flags.EnDagaRusa).Posicion.X, .Usuarios(UserList(ID).flags.EnDagaRusa).Posicion.Y, True
        .Usuarios(UserList(ID).flags.EnDagaRusa).ID = 0
        UserList(ID).flags.EnDagaRusa = 0
        .Cupos = .Cupos + 1
        WritePauseToggle ID
    End With

End Sub

Private Sub Limpiar()

    With DagaRusa
        .Activo = False
        .Conteo = -1
        .Cupos = 0
        .Inscripcion = 0
        .Premio = 0
        .Restantes = 0
        .Total = 0
        .AtacoUser = 0
        .Atacar = 0
        .PuedeAtacar = False
        .ActivarEvento = False
        Erase .Usuarios()
    End With
    INDEX_NPC_DAGA_RUSA_ONLINE = 0
End Sub

Private Function Tiene_Objeto(ByVal ID As Integer) As Boolean
    Dim LoopC As Long
    Tiene_Objeto = False
    With UserList(ID)
        For LoopC = 1 To .CurrentInventorySlots
            If .Invent.Object(LoopC).ObjIndex > 0 Then Exit Function
        Next LoopC
        Tiene_Objeto = True
    End With
End Function

