Attribute VB_Name = "Admin"
'Argentum Online 0.12.2
'Copyright (C) 2002 Marquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

Public Type tMotd

    texto As String
    Formato As String

End Type

Public MaxLines As Integer

Public MOTD()   As tMotd

Public Type tAPuestas

    Ganancias As Long
    Perdidas As Long
    Jugadas As Long

End Type

Public Apuestas                          As tAPuestas

Public tInicioServer                     As Long


'INTERVALOS
Public SanaIntervaloSinDescansar         As Integer

Public StaminaIntervaloSinDescansar      As Integer

Public SanaIntervaloDescansar            As Integer

Public StaminaIntervaloDescansar         As Integer

Public StaminaIntervaloLloviendo         As Integer

Public IntervaloSed                      As Integer

Public IntervaloHambre                   As Integer

Public IntervaloVeneno                   As Integer

Public IntervaloIncinerado               As Integer

Public IntervaloParalizado               As Integer

Public Const IntervaloParalizadoReducido As Integer = 37

Public IntervaloInvisible                As Integer

Public IntervaloFrio                     As Integer

Public IntervaloWavFx                    As Integer

Public IntervaloLanzaHechizo             As Integer

Public IntervaloNPCPuedeAtacar           As Integer

Public IntervaloInvocacion               As Integer

Public IntervaloOculto                   As Integer '[Nacho]

Public IntervaloUserPuedeAtacar          As Long

Public IntervaloGolpeUsar                As Long

Public IntervaloMagiaGolpe               As Long

Public IntervaloGolpeMagia               As Long

Public IntervaloUserPuedeCastear         As Long

Public IntervaloUserPuedeTrabajar        As Long

Public IntervaloParaConexion             As Long

Public IntervaloCerrarConexion           As Long '[Gonzalo]

Public IntervaloUserPuedeUsar            As Long

Public IntervaloFlechasCazadores         As Long

Public IntervaloPuedeMakrear             As Integer

Public IntervaloPuedeSerAtacado          As Long

Public IntervaloAtacable                 As Long

Public IntervaloOwnedNpc                 As Long

Public IntervaloOcultable                As Long

Public IntervaloTocar                    As Long

Public Intervalo_Global                  As Long

Public IntervaloCaminar                  As Long

'BALANCE

Public PorcentajeRecuperoMana            As Integer

Public MinutosWs                         As Long

Public MinutosGuardarUsuarios            As Long

Public IntervaloReconexionDB             As Long

Public Puerto                            As Integer

Public BootDelBackUp                     As Boolean

Public Lloviendo                         As Boolean

Public DificultadExtraer                 As Integer


Function VersionOK(ByVal Ver As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    VersionOK = (Ver = ULTIMAVERSION)

End Function

Sub ReSpawnOrigPosNpcs()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    If frmMain.Visible Then frmMain.txtStatus.Text = "Haciendo ReSpawn de NPCS en posicion original"

    Dim i     As Integer

    Dim MiNPC As NPC
       
    For i = 1 To LastNPC

        'OJO
        If Npclist(i).flags.NPCActive Then
            
            If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)

            End If
            
            'tildada por sugerencia de yind
            'If Npclist(i).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(i, 0)
            'End If
        End If
       
    Next i
    
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Respawn NPCS en posicion original finalizado."

End Sub

Sub WorldSave()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim loopX As Integer

    Dim hFile As Integer
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER))
    
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    
    Dim j As Integer, K As Integer
    
    For j = 1 To NumMaps

        If MapInfo(j).BackUp = 1 Then K = K + 1
    Next j
    
    FrmStat.ProgressBar1.min = 0
    FrmStat.ProgressBar1.max = K
    FrmStat.ProgressBar1.Value = 0
    
    For loopX = 1 To NumMaps
        'DoEvents
        
        If MapInfo(loopX).BackUp = 1 Then
            Call GrabarMapa(loopX, App.Path & "\WorldBackUp\Mapa" & loopX)
            FrmStat.ProgressBar1.Value = FrmStat.ProgressBar1.Value + 1

        End If
    
    Next loopX
    
    FrmStat.Visible = False
    
    If FileExist(DatPath & "\bkNpcs.dat") Then Kill (DatPath & "bkNpcs.dat")
    
    hFile = FreeFile()
    
    Open DatPath & "\bkNpcs.dat" For Output As hFile
    
    For loopX = 1 To LastNPC

        If Npclist(loopX).flags.BackUp = 1 Then
            Call BackUPnPc(loopX, hFile)

        End If

    Next loopX
        
    Close hFile
    
    Call SaveForums
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluido.", FontTypeNames.FONTTYPE_SERVER))

End Sub

Public Sub Encarcelar(ByVal UserIndex As Integer, _
                      ByVal Minutos As Long, _
                      Optional ByVal GmName As String = vbNullString)
    '***************************************************
    'Author: Lucas Recoaro
    'Last Modification: 26/08/2018
    'Shak: Agregamos el array.
    'Recox: Arreglado problema de tiempo en carcel
    '***************************************************
    
    With UserList(UserIndex)
    
        'Si esta navegando, lo bajamos de la barca
        If .flags.Navegando = 1 Then
            Call DejardeNavegar(UserIndex)
        End If
        
        'Si esta equitando, lo bajamos de la montura
        If .flags.Equitando = 1 Then
            Call UnmountMontura(UserIndex)
            Call WriteEquitandoToggle(UserIndex)
        End If
    
        .Counters.Pena = Minutos * 60
        
        Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
        
        If LenB(GmName) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberas permanecer en la carcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberas permanecer en la carcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    
        End If
    
    End With

End Sub

Public Function BorrarUsuario(ByVal UserIndex As Integer, ByVal Slot As Byte) As Boolean

    '********************************************************************************
    'Author: Lorwik
    'Last Modification: 21/05/2020
    '********************************************************************************
    
    Dim username As String
    
    'Obtenemos el nombre del usuario
    username = UserList(UserIndex).AccountInfo.AccountPJ(Slot).name
    
    'Podria estar de mas, pero... Existe el personaje?
    If Not PersonajeExiste(username) Then
        BorrarUsuario = False
        Exit Function
    End If
    
    'Mandamos "borrar" en la base de datos (en realidad no se borra)
    Call BorrarUsuarioDatabase(username)
    
    'Actualizamo las listas y el cliente del usuario
    Call DeletePJCuenta(UserIndex, Slot)
    
    BorrarUsuario = True

End Function

Public Function BANCheck(ByVal name As String) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 18/09/2018
    '18/09/2018 CHOTS: Checks database too
    '***************************************************

    BANCheck = BANCheckDatabase(name)

End Function

Public Function PersonajeExiste(ByVal username As String) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 18/09/2018
    '18/09/2018 CHOTS: Checks database too
    '***************************************************

    PersonajeExiste = PersonajeExisteDatabase(username)

End Function

Public Function CuentaExiste(ByVal username As String) As Boolean

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 12/10/2018
    '***************************************************

    CuentaExiste = CuentaExisteDatabase(username)

End Function

Public Sub UnBan(ByVal name As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 18/09/2018
    '18/09/2018 CHOTS: Checks database too
    '***************************************************
    
    Call UnBanDatabase(name)

    'Remove it from the banned people database
    Call WriteVar(App.Path & "\Dat\" & "BanDetail.dat", name, "BannedBy", "NOBODY")
    Call WriteVar(App.Path & "\Dat\" & "BanDetail.dat", name, "Reason", "NO REASON")

End Sub

Public Function GetUserGuildIndex(ByVal username As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso
    'Last Modification: 18/09/2018
    '18/09/2018 CHOTS: Checks database too
    '***************************************************
    If InStrB(username, "\") <> 0 Then
        username = Replace(username, "\", vbNullString)

    End If

    If InStrB(username, "/") <> 0 Then
        username = Replace(username, "/", vbNullString)

    End If

    If InStrB(username, ".") <> 0 Then
        username = Replace(username, ".", vbNullString)

    End If

    GetUserGuildIndex = GetUserGuildIndexDatabase(username)

End Function

Public Sub CopyUser(ByVal username As String, ByVal newName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 18/09/2018
    '18/09/2018 CHOTS: Checks database too
    '***************************************************
    
    Call CopyUserDatabase(username, newName)

End Sub

Public Sub BanIpAgrega(ByVal IP As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call BanIps.Add(IP)
    Call BanIpGuardar
    
    ' Agrego la regla al firewall para que bloquee la IP
    Call Shell("netsh.exe advfirewall firewall add rule name=""Baneo de IP " & IP & """ dir=in protocol=any action=block remoteip=" & IP)
    
End Sub

Public Function BanIpBuscar(ByVal IP As String) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Dale  As Boolean
    Dim LoopC As Long
    
    Dale = True
    LoopC = 1

    Do While LoopC <= BanIps.Count And Dale
        Dale = (BanIps.Item(LoopC) <> IP)
        LoopC = LoopC + 1
    Loop
    
    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = LoopC - 1

    End If

End Function

Public Function BanIpQuita(ByVal IP As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim n As Long
    
    n = BanIpBuscar(IP)

    If n > 0 Then
        Call BanIps.Remove(n)
        Call BanIpGuardar
        
        ' Agrego la regla al firewall para que borre la regla de la IP a desbanear.
        Call Shell("netsh.exe advfirewall firewall delete rule name=""Baneo de IP " & IP & """ dir=in protocol=any action=block remoteip=" & IP)
        
        BanIpQuita = True
    Else
        BanIpQuita = False

    End If

End Function

Public Sub BanIpGuardar()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ArchivoBanIp As String

    Dim ArchN        As Long

    Dim LoopC        As Long
    
    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN
    
    For LoopC = 1 To BanIps.Count
        Print #ArchN, BanIps.Item(LoopC)
    Next LoopC
    
    Close #ArchN

End Sub

Public Sub BanIpCargar()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ArchN        As Long

    Dim Tmp          As String

    Dim ArchivoBanIp As String
    
    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    Set BanIps = New Collection
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN
    
    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanIps.Add Tmp
    Loop
    
    Close #ArchN

End Sub

Public Function UserDarPrivilegioLevel(ByVal name As String) As PlayerType
    '***************************************************
    'Author: Unknown
    'Last Modification: 03/02/07
    'Last Modified By: Juan Martin Sotuyo Dodero (Maraxus)
    '***************************************************

    If EsAdmin(name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.User

    End If

End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, _
                        ByVal username As String, _
                        ByVal Reason As String)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 03/02/07
    '22/05/2010: Ya no se peude banear admins de mayor rango si estan online.
    '***************************************************

    Dim tUser     As Integer

    Dim UserPriv  As Byte

    Dim cantPenas As Byte

    Dim rank      As Integer
    
    If InStrB(username, "+") Then
        username = Replace(username, "+", " ")

    End If
    
    tUser = NameIndex(username)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)

        If tUser <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_SERVER)
            
            If PersonajeExiste(username) Then
                UserPriv = UserDarPrivilegioLevel(username)
                
                If (UserPriv And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If BANCheck(username) Then
                        Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call LogBanFromName(username, bannerUserIndex, Reason)
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado a " & username & ".", FontTypeNames.FONTTYPE_SERVER))
                        
                        Call SaveBan(username, Reason, .name)
                        
                        If (UserPriv And rank) = (.flags.Privilegios And rank) Then
                            .flags.Ban = 1
                            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                            Call CloseUser(bannerUserIndex)

                        End If
                        
                        Call LogGM(.name, "BAN a " & username)

                    End If

                End If

            Else
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & username & " no existe.", FontTypeNames.FONTTYPE_INFO)

            End If

        Else

            If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
            Else
            
                Call LogBan(tUser, bannerUserIndex, Reason)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado a " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_SERVER))
                
                'Ponemos el flag de ban a 1
                UserList(tUser).flags.Ban = 1
                
                If (UserList(tUser).flags.Privilegios And rank) = (.flags.Privilegios And rank) Then
                    .flags.Ban = 1
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                    Call CloseUser(bannerUserIndex)

                End If
                
                Call LogGM(.name, "BAN a " & username)
                
                Call SaveBan(username, Reason, .name)
                
                Call CloseUser(tUser)

            End If

        End If

    End With

End Sub

