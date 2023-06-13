Attribute VB_Name = "modTrampas"
Option Explicit

Private Const MAP_MANSION    As Byte = 66

Private Const CasaSpiritsORO As Integer = 30000

Private Const MansionRayoX1 As Byte = 47

Private Const MansionRayoX2 As Byte = 51

Private Const MansionRayoY As Byte = 22

Private Const FX_MANSION     As Byte = 10

Sub VigilarEventosTrampas(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Revisión: Edurne
    'Last Modification: 10/09/2015
    '***************************************************
    With UserList(UserIndex).Pos

        'Esta en casa encantada?
        If .Map = MAP_MANSION Then Call CasaEncantada(UserIndex)
        
        'Esta en alguno de estos Triggers?
        If MapData(.Map, .X, .Y).Trigger = eTrigger.TRAMPA_1 Then
            Call Trampa(UserIndex, 34)
            
        ElseIf MapData(.Map, .X, .Y).Trigger = eTrigger.TRAMPA_2 Then
            Call Trampa(UserIndex, 37)

        End If

    End With

End Sub

Public Sub Trampa(ByVal UserIndex As Integer, Tipotrampa As Integer)

    On Error GoTo fallo

    Dim daño As Integer
    
    'Calculamos el daño de la trampa
    daño = RandomNumber(5, 20)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Tipotrampa, 0))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(218, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - daño
    Call WriteConsoleMsg(UserIndex, "¡¡Una trampa te causa " & daño & " de daño!!", FontTypeNames.FONTTYPE_FIGHT)
    Call WriteUpdateHP(UserIndex)

    If UserList(UserIndex).Stats.MinHp <= 0 Then Call UserDie(UserIndex)
    Exit Sub
fallo:
    Call LogError("TRAMPA " & Err.Number & " D: " & Err.description)

End Sub

Sub CasaEncantada(ByVal UserIndex As Integer)

    'Creado por Pluto, adaptado y mejorado por Lorwik
    'pluto:2.17
    Dim X         As Byte

    Dim Y         As Byte

    Dim Map       As Integer

    Dim DadosCasa As Byte

    With UserList(UserIndex)
        Map = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y
        
        'pluto:rayos puerta
        If (X = MansionRayoX1 Or X = MansionRayoX2) And Y = MansionRayoY And .flags.Muerto = 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(113, X, Y))
        
        If .Counters.Morph > 0 Then Exit Sub
        
        'pluto:sala sangre casa
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.SALASANGRE Then
            Call WriteConsoleMsg(UserIndex, "¡¡La habitación de sangre te ha matado!!", FontTypeNames.FONTTYPE_FIGHT)
            Call UserDie(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(102, .Pos.X, .Pos.Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FX_MANSION, 0))
            Exit Sub

        End If
        
        'Lorwik:Espitirus
        DadosCasa = RandomNumber(1, 300)

        Select Case DadosCasa
        
            Case 1

                If .Stats.Gld >= 3000 Then
                    Call WriteConsoleMsg(UserIndex, "Los Espiritus de la mansión te hacen perder Oro.", FontTypeNames.FONTTYPE_FIGHT)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(102, .Pos.X, .Pos.Y))
                    Call TirarOro(3000, UserIndex)
                    Call WriteUpdateUserStats(UserIndex)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FX_MANSION, 0))
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(.Char.CharIndex, 3, 0))
                    Exit Sub

                End If
            
            Case 30

                If .flags.Morph = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Los Espiritus de la Casa te transforman en Cerdo.", FontTypeNames.FONTTYPE_FIGHT)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(102, .Pos.X, .Pos.Y))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FX_MANSION, 0))
                    .flags.Morph = .Char.body
                    .Counters.Morph = IntervaloMorphPJ
                    Call ChangeUserChar(UserIndex, 6, 0, UserList(UserIndex).Char.Heading, 2, 2, 2, 0, 0)
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(.Char.CharIndex, 3, 0))
                    Exit Sub

                End If
            
            Case 49
                Call WriteConsoleMsg(UserIndex, "Los Espiritus de la mansión te hacen perder el inventario.", FontTypeNames.FONTTYPE_FIGHT)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(102, .Pos.X, .Pos.Y))
                Call TirarTodosLosItems(UserIndex)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FX_MANSION, 0))
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(.Char.CharIndex, 3, 0))
                Exit Sub
            
            Case 53
                Call WriteConsoleMsg(UserIndex, "Los Espiritus de la mansión te teleportan fuera de ella.", FontTypeNames.FONTTYPE_FIGHT)
                Call PrepareMessagePlayWave(102, 0, 0)
                Call WarpUserChar(UserIndex, MAP_MANSION, 50, 69, True)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FX_MANSION, 0))
                Exit Sub
            
            Case 97
                Call WriteConsoleMsg(UserIndex, "Los Espiritus de la mansión te han Paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Call PrepareMessagePlayWave(102, 0, 0)
                .flags.Paralizado = 1
                .Counters.Paralisis = IntervaloParalizado
                Call WriteParalizeOK(UserIndex)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(247, .Pos.X, .Pos.Y))
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 8, 0))
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(.Char.CharIndex, 3, 0))
                Exit Sub

        End Select

    End With

End Sub

