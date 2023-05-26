Attribute VB_Name = "modPLantes"
Option Explicit

Private Const MAPA_PLANTES As Byte = 61
Private Const X1 As Byte = 63
Private Const Y1 As Byte = 16
Private Const X2 As Byte = 64
Private Const Y2 As Byte = 16

Private Const MIN_LEVEL As Byte = 25

Public Oponente(0 To 1) As Byte

Private PlantesenCurso As Boolean

Public Sub InscripcionPlantes(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 26/05/2023
'***************************************************

    With UserList(UserIndex)
    
        If Not PuedePlantar(UserIndex) Then Exit Sub
    
        '¿Hay alguien esperando duelo?
        If Oponente(0) = 0 Then
            'No lo hay, pues lo metemos en la cola y le asignamos el puesto 0
            Oponente(0) = UserIndex
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Plantes: " & UserList(UserIndex).name & " está buscando contrincante.", FontTypeNames.FONTTYPE_DIOS))
        Else
            
            '¿El contrincante esta en otro evento?
            If UserList(Oponente(0)).flags.EstaDueleando Or UserList(Oponente(0)).flags.ArenaRinkel Then
                Call WriteConsoleMsg(Oponente(1), "El contrincante esta en otro evento. El duelo se ha cancelado.", FontTypeNames.FONTTYPE_INFO)
                Oponente(0) = 0
                Oponente(1) = 0
            End If
        
            'Si lo hay, le asignamos el puesto 1 y para dentro.
            Oponente(1) = UserIndex
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Plantes: ¡" & UserList(Oponente(1)).name & " aceptó el desafío!", FontTypeNames.FONTTYPE_DIOS))
            
            Call ComenzarDuelo(Oponente(0), Oponente(1))

        End If
        
    End With

End Sub

Private Function PuedePlantar(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Lorwik
'Fecha: 26/05/2023
'***************************************************

    With UserList(UserIndex)
    
        If .Stats.ELV < MIN_LEVEL Then
            Call WriteConsoleMsg(UserIndex, "Tu nivel debe de ser " & MIN_LEVEL & " o superior.", FontTypeNames.FONTTYPE_INFO)
            PuedePlantar = False
            Exit Function

        End If
        
        If .flags.Muerto = 1 Then '¿Esta Muerto?
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            PuedePlantar = False
            Exit Function

        End If
        
        '¿Está trabajando?
        If .flags.MacroTrabajo <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            PuedePlantar = False
            Exit Function

        End If
        
        If MapInfo(.Pos.Map).Pk Then '¿esta en zona insegura?
            Call WriteConsoleMsg(UserIndex, "¡¡Tienes que estas en zona segura!!", FontTypeNames.FONTTYPE_INFO)
            PuedePlantar = False
            Exit Function

        End If
        
        If PlantesenCurso Then
            Call WriteConsoleMsg(UserIndex, "Hay un duelo de plantes en curso, espera a que termine.", FontTypeNames.FONTTYPE_INFO)
            PuedePlantar = False
            Exit Function

        End If
        
        If .flags.EstaDueleando Then
            Call WriteConsoleMsg(UserIndex, "Ya estas en duelos ranked.", FontTypeNames.FONTTYPE_INFO)
            PuedePlantar = False
            Exit Function

        End If
        
        If .flags.ArenaRinkel = True Then
            Call WriteConsoleMsg(UserIndex, "Ya estas en la arena de rinkel", FontTypeNames.FONTTYPE_INFO)
            PuedePlantar = False
            Exit Function
        End If
        
        If .flags.EstaPlantando Then
            Call WriteConsoleMsg(UserIndex, "Ya estas en duelos de plantes.", FontTypeNames.FONTTYPE_INFO)
            PuedePlantar = False
            Exit Function

        End If
    
    End With
    
    PuedePlantar = True
End Function

Public Sub ComenzarDuelo(ByVal UserIndex As Integer, ByVal tIndex As Integer)
    Oponente(0) = 0
    Oponente(1) = 0

    'Sacamos escudos
    With UserList(UserIndex)
        If .Invent.EscudoEqpSlot > 0 Then _
            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
        
        .flags.EstaPlantando = True
    End With
    
    With UserList(tIndex)
        If .Invent.EscudoEqpSlot > 0 Then _
            Call Desequipar(tIndex, .Invent.EscudoEqpSlot)
        
        .flags.EstaPlantando = True
    End With
    
    'Paralizamos a los contricantes
    UserList(UserIndex).flags.Paralizado = 1
    UserList(tIndex).flags.Paralizado = 1
    
    UserList(UserIndex).flags.Oponente = tIndex
    UserList(tIndex).flags.Oponente = UserIndex
    
    'Los llevamos al mapa de duelo
    Call WarpUserChar(UserIndex, MAPA_PLANTES, X1, Y1, True)
    Call WarpUserChar(tIndex, MAPA_PLANTES, X2, Y2, True)
    
    PlantesenCurso = True

End Sub

Public Sub TerminarPlantes(ByVal Ganador As Integer, ByVal Perdedor As Integer)

    Dim ELOGANADOR  As Long

    Dim ELOPERDEDOR As Long

    UserList(Ganador).Stats.MinMAN = UserList(Ganador).Stats.MaxMAN
    UserList(Ganador).Stats.MinHp = UserList(Ganador).Stats.MaxHp
    UserList(Perdedor).Stats.MinMAN = UserList(Perdedor).Stats.MaxMAN
    UserList(Perdedor).Stats.MinHp = UserList(Perdedor).Stats.MaxHp
    
    Call WriteUpdateUserStats(Ganador)
    Call WriteUpdateUserStats(Perdedor)

    With UserList(Perdedor)

        UserList(Ganador).flags.EsperandoDuelo = False
        UserList(Ganador).flags.Oponente = 0
        UserList(Ganador).flags.EstaDueleando = False
        UserList(Ganador).flags.TimeDuelo = 9
        UserList(Ganador).flags.GanoDuelo = True
            
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Plantes: ¡" & UserList(Ganador).flags.PerdioRonda & "-" & UserList(Perdedor).flags.PerdioRonda & " para " & UserList(Ganador).name & "! ¡" & UserList(Ganador).name & " Gana!", FontTypeNames.FONTTYPE_DIOS))
            
        'Calcularmos el ELO
        ELOGANADOR = CalcularELO(Ganador, Perdedor, True)
        ELOPERDEDOR = CalcularELO(Perdedor, Ganador, False)
        'Lo asignamos
        UserList(Ganador).Stats.ELO = ELOGANADOR + UserList(Ganador).Stats.ELO
        Call WriteConsoleMsg(Ganador, "Plantes: ¡Has ganado +" & ELOGANADOR & " puntos! Tu ELO actual es de " & UserList(Ganador).Stats.ELO & ".", FontTypeNames.FONTTYPE_INFOBOLD)
        UserList(Perdedor).Stats.ELO = ELOPERDEDOR + UserList(Perdedor).Stats.ELO
        Call WriteConsoleMsg(Perdedor, "Plantes: ¡Has perdido " & ELOPERDEDOR & " puntos! Tu ELO actual es de " & UserList(Perdedor).Stats.ELO & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            
        Call resetPlantes(Ganador, Perdedor)
        
        Call SaveUser(Perdedor)
        Call SaveUser(Ganador)
        Call ActualizarRank(Ganador)
        Call ActualizarRank(Perdedor)
            
        PlantesenCurso = False

    End With

End Sub

Public Sub resetPlantes(ByVal Ganador As Integer, ByVal Perdedor As Integer)

    'Reseteamos los flags del Ganador
    With UserList(Ganador)
        .flags.EsperandoDuelo = False
        .flags.Oponente = 0
        .flags.EstaPlantando = False
    End With

    'Reseteamos los Flags Perdedor
    With UserList(Perdedor)
        .flags.EsperandoDuelo = False
        .flags.Oponente = 0
        .flags.EstaPlantando = False
    End With
    
    'Teletransportamso a su casa
    Call MandaraCasa(Ganador)
    Call MandaraCasa(Perdedor)
    
    Call RemoveParalisis(Ganador)
    Call RemoveParalisis(Perdedor)

End Sub
