Attribute VB_Name = "modDuelos"
Option Explicit

Private Const MapaDuelos     As Byte = 61
Private Const XEsquinaAbajo  As Byte = 52
Private Const YEsquinaAbajo  As Byte = 68
Private Const XEsquinaArriba As Byte = 30
Private Const YEsquinaArriba As Byte = 47

Private Const MIN_LEVEL As Byte = 25

Private DuelosenCurso As Boolean

Public Oponente(0 To 1)      As Byte
    
Public Sub resetDueloSet(ByVal Ganador As Integer, ByVal Perdedor As Integer)

    'Reseteamos los flags del Ganador
    With UserList(Ganador)
        .flags.EsperandoDuelo = False
        .flags.Oponente = 0
        .flags.EstaDueleando = False
        .flags.PerdioRonda = 0
    End With

    'Reseteamos los Flags Perdedor
    With UserList(Perdedor)
        .flags.EsperandoDuelo = False
        .flags.Oponente = 0
        .flags.EstaDueleando = False
        .flags.PerdioRonda = 0
    End With
    
    'Teletransportamso a su casa
    Call MandaraCasa(Ganador)
    Call MandaraCasa(Perdedor)

End Sub

Public Sub EsperarOponenteDuelo(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)

        If .Stats.ELV < MIN_LEVEL Then
            Call WriteConsoleMsg(UserIndex, "Tu nivel debe de ser " & MIN_LEVEL & " o superior.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .flags.Muerto = 1 Then '¿Esta Muerto?
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        '¿Está trabajando?
        If .flags.MacroTrabajo <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If MapInfo(.Pos.Map).Pk Then '¿esta en zona insegura?
            Call WriteConsoleMsg(UserIndex, "¡¡Tienes que estas en zona segura!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If DuelosenCurso Then '¿Hay gente duelando?
            Call WriteConsoleMsg(UserIndex, "Hay un duelo en curso, espera a que termine.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .flags.EstaPlantando Then
            Call WriteConsoleMsg(UserIndex, "Ya estas en duelos de plantes.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If .flags.ArenaRinkel = True Then
            Call WriteConsoleMsg(UserIndex, "Ya estas en la arena de rinkel", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.EstaDueleando Then
            Call WriteConsoleMsg(UserIndex, "Ya estas en la cola de duelos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        '¿Hay alguien esperando duelo?
        If Oponente(0) = 0 Then
            'No lo hay, pues lo metemos en la cola y le asignamos el puesto 0
            Oponente(0) = UserIndex
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranked: " & UserList(UserIndex).name & " está buscando contrincante.", FontTypeNames.FONTTYPE_DIOS))
        Else
        
            '¿El contrincante esta en otro evento?
            If UserList(Oponente(0)).flags.EstaPlantando Or UserList(Oponente(0)).flags.ArenaRinkel Then
                Call WriteConsoleMsg(Oponente(1), "El contrincante esta en otro evento. El duelo se ha cancelado.", FontTypeNames.FONTTYPE_INFO)
                Oponente(0) = 0
                Oponente(1) = 0
            End If
        
            'Si lo hay, le asignamos el puesto 1 y para dentro.
            Oponente(1) = UserIndex
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranked: ¡" & UserList(Oponente(1)).name & " aceptó el desafío!", FontTypeNames.FONTTYPE_DIOS))
            
            Call ComenzarDuelo(Oponente(0), Oponente(1))

        End If

    End With

End Sub

Public Sub ComenzarDuelo(ByVal UserIndex As Integer, ByVal tIndex As Integer)
    Oponente(0) = 0
    Oponente(1) = 0
    UserList(UserIndex).flags.EstaDueleando = True
    UserList(UserIndex).flags.Oponente = tIndex
    UserList(UserIndex).flags.PerdioRonda = 0
    
    Call WarpUserChar(UserIndex, MapaDuelos, XEsquinaAbajo, YEsquinaAbajo, True) 'esqina de duelos
    UserList(tIndex).flags.EstaDueleando = True
    UserList(tIndex).flags.Oponente = UserIndex
    UserList(tIndex).flags.PerdioRonda = 0
    
    Call WarpUserChar(tIndex, MapaDuelos, XEsquinaArriba, YEsquinaArriba, True) 'esqina de duelos
    
    DuelosenCurso = True

End Sub
   
Public Sub TerminarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)

    Dim ELOGANADOR  As Long

    Dim ELOPERDEDOR As Long

    '28/10/2015 Irongete: Le subo la vida y el maná a los dos jugadores
    UserList(Ganador).Stats.MinMAN = UserList(Ganador).Stats.MaxMAN
    UserList(Ganador).Stats.MinHp = UserList(Ganador).Stats.MaxHp
    UserList(Perdedor).Stats.MinMAN = UserList(Perdedor).Stats.MaxMAN
    UserList(Perdedor).Stats.MinHp = UserList(Perdedor).Stats.MaxHp
    Call WriteUpdateUserStats(Ganador)
    Call WriteUpdateUserStats(Perdedor)

    With UserList(Perdedor)

        If .flags.PerdioRonda = 1 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranked: ¡" & UserList(Ganador).flags.PerdioRonda & "-" & UserList(Perdedor).flags.PerdioRonda & " para " & UserList(Ganador).name & "!", FontTypeNames.FONTTYPE_DIOS))
            Call WarpUserChar(Perdedor, MapaDuelos, XEsquinaAbajo, YEsquinaAbajo, True) 'esqina de duelos
            Call WarpUserChar(Ganador, MapaDuelos, XEsquinaArriba, YEsquinaArriba, True) 'esqina de duelos
            
        ElseIf .flags.PerdioRonda >= 2 Then
            UserList(Ganador).flags.EsperandoDuelo = False
            UserList(Ganador).flags.Oponente = 0
            UserList(Ganador).flags.EstaDueleando = False
            UserList(Ganador).flags.TimeDuelo = 9
            UserList(Ganador).flags.GanoDuelo = True
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranked: ¡" & UserList(Ganador).flags.PerdioRonda & "-" & UserList(Perdedor).flags.PerdioRonda & " para " & UserList(Ganador).name & "! ¡" & UserList(Ganador).name & " Gana!", FontTypeNames.FONTTYPE_DIOS))
            
            'Calcularmos el ELO
            ELOGANADOR = CalcularELO(Ganador, Perdedor, True)
            ELOPERDEDOR = CalcularELO(Perdedor, Ganador, False)
            'Lo asignamos
            UserList(Ganador).Stats.ELO = ELOGANADOR + UserList(Ganador).Stats.ELO
            Call WriteConsoleMsg(Ganador, "Ranked: ¡Has ganado +" & ELOGANADOR & " puntos! Tu ELO actual es de " & UserList(Ganador).Stats.ELO & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            UserList(Perdedor).Stats.ELO = ELOPERDEDOR + UserList(Perdedor).Stats.ELO
            Call WriteConsoleMsg(Perdedor, "Ranked: ¡Has perdido " & ELOPERDEDOR & " puntos! Tu ELO actual es de " & UserList(Perdedor).Stats.ELO & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            
            Call resetDueloSet(Ganador, Perdedor)
        
            Call SaveUser(Perdedor)
            Call SaveUser(Ganador)
            Call ActualizarRank(Ganador)
            Call ActualizarRank(Perdedor)
            
            DuelosenCurso = False

        End If

    End With

End Sub
