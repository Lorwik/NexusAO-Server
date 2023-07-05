Attribute VB_Name = "Trabajo"
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

Private Const GASTO_ENERGIA As Byte = 6

Public Const ESFUERZOEXTRAER As Byte = 3

Private Const PRECIOINSTRUCCION As Long = 50000

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)

    '********************************************************
    'Autor: Nacho (Integer)
    'Last Modif: 11/19/2009
    'Chequea si ya debe mostrarse
    'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
    '********************************************************
    On Error GoTo errHandler

    With UserList(UserIndex)
        .Counters.TiempoOculto = .Counters.TiempoOculto - 1

        If .Counters.TiempoOculto <= 0 Then
            If .clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
                If .Invent.ArmourEqpObjIndex = 648 Or .Invent.ArmourEqpObjIndex = 360 Then
                    .Counters.TiempoOculto = IntervaloOculto
                    Exit Sub

                End If

            End If

            .Counters.TiempoOculto = 0
            .flags.Oculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Mercenario Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SetInvisible(UserIndex, .Char.CharIndex, False)

                End If

            End If

        End If

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en Sub DoPermanecerOculto")

End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 13/01/2010 (ZaMa)
    'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
    'Modifique la formula y ahora anda bien.
    '13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
    '***************************************************

    On Error GoTo errHandler

    Dim Suerte As Double

    Dim res    As Integer

    Dim Skill  As Integer
    
    With UserList(UserIndex)
        Skill = .Stats.UserSkills(eSkill.Ocultarse)
        
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        res = RandomNumber(1, 100)
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            Suerte = (-0.000001 * (100 - Skill) ^ 3)
            Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            Suerte = Suerte + (-0.0088 * (100 - Skill))
            Suerte = Suerte + (0.9571)
            Suerte = Suerte * IntervaloOculto
            
            If .clase = eClass.Bandit Then
                .Counters.TiempoOculto = Int(Suerte / 2)
            Else
                .Counters.TiempoOculto = Suerte

            End If
            
            ' No es pirata o es uno sin barca
            If .flags.Navegando = 0 Then
                Call SetInvisible(UserIndex, .Char.CharIndex, True)
        
                Call WriteConsoleMsg(UserIndex, "Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
                ' Es un pirata navegando
            Else
                ' Le cambiamos el body a galeon fantasmal
                .Char.body = iFragataFantasmal
                ' Actualizamos clientes
                Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

            End If
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 4 Then
                Call WriteConsoleMsg(UserIndex, "No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 4

            End If

            '[/CDT]
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, False)

        End If
        
        .Counters.Ocultando = .Counters.Ocultando + 1

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, _
                    ByRef Barco As ObjData, _
                    ByVal Slot As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2020 (Recox)
'13/01/2010: ZaMa - El pirata pierde el ocultar si desequipa barca.
'16/09/2010: ZaMa - Ahora siempre se va el invi para los clientes al equipar la barca (Evita cortes de cabeza).
'10/12/2010: Pato - Limpio las variables del inventario que hacen referencia a la barca, sino el pirata que la ultima barca que equipo era el galeon no explotaba(Y capaz no la tenia equipada :P).
'12/01/2020: Recox - Se refactorizo un poco para reutilizar con monturas .
'***************************************************
    
    With UserList(UserIndex)
        If .flags.Equitando = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes navegar mientras estas en tu montura!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        '�Es una montura acuatica? Pedimo Skills en Equitacion
'        If Barco.MontTipo = 1 Then
'            If UserList(UserIndex).Stats.UserSkills(Equitacion) < Barco.MinSkill Then
'                Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Barco.MinSkill & " puntos en equitaci�n.", FontTypeNames.FONTTYPE_INFO)
'                Exit Sub
'            End If
'
'        Else
            
            If .Stats.UserSkills(eSkill.Navegacion) < Barco.MinSkill Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
    
            End If
        
        'End If
        
        ' No estaba navegando
        If .flags.Navegando = 0 Then
            
            Call ComenzaraNavegar(UserIndex, Slot)
        
        ' Estaba navegando
        Else
        
            Call DejardeNavegar(UserIndex)

        End If

    End With
    
End Sub

Public Sub ComenzaraNavegar(ByVal UserIndex As Integer, ByVal Slot As Integer)

    With UserList(UserIndex)
    
        .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
        .Invent.BarcoSlot = Slot
            
        .Char.Head = 0
            
        ' No esta muerto
        If .flags.Muerto = 0 Then
            Call ToggleBoatBody(UserIndex)
            Call SetVisibleStateForUserAfterNavigateOrEquitate(UserIndex)
                
        ' Esta muerto
        Else
            .Char.body = iFragataFantasmal
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .Char.AuraAnim = NingunAura
            .Char.AuraColor = NingunAura
                
        End If
            
        ' Comienza a navegar
        .flags.Navegando = 1
        
        ' Actualizo clientes
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
    
        Call WriteNavigateToggle(UserIndex)
    
    End With

End Sub

Public Sub DejardeNavegar(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
        .Invent.BarcoObjIndex = 0
        .Invent.BarcoSlot = 0
    
        ' No esta muerto
        If .flags.Muerto = 0 Then
            .Char.Head = .OrigChar.Head
                
            Call SetEquipmentOnCharAfterNavigateOrEquitate(UserIndex)
                
            ' Al dejar de navegar, si estaba invisible actualizo los clientes
            If .flags.invisible = 1 Then
                Call SetInvisible(UserIndex, .Char.CharIndex, True)
            End If
                
        ' Esta muerto
        Else
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .Char.AuraAnim = NingunAura
            .Char.AuraColor = NingunAura

        End If
            
        ' Termina de navegar
        .flags.Navegando = 0
        
        ' Actualizo clientes
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
    
        Call WriteNavigateToggle(UserIndex)
    
    End With
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    With UserList(UserIndex)

        If .flags.TargetObjInvIndex > 0 Then
           
            If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(.flags.TargetObjInvIndex).MinSkill <= .Stats.UserSkills(eSkill.Mineria) Then
                Call DoLingotes(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de mineria suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)

            End If
        
        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.description)

End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, _
                      ByVal cant As Long, _
                      ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/07/2010
    '10/07/2010: ZaMa - Ahora cant es long para evitar un overflow.
    '***************************************************

    Dim i     As Integer

    Dim Total As Long

    For i = 1 To UserList(UserIndex).CurrentInventorySlots

        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Total = Total + UserList(UserIndex).Invent.Object(i).Amount

        End If

    Next i
    
    If cant <= Total Then
        TieneObjetos = True
        Exit Function

    End If
        
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, _
                         ByVal cant As Integer, _
                         ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 05/08/09
    '05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
    '***************************************************

    Dim i As Integer

    For i = 1 To UserList(UserIndex).CurrentInventorySlots

        With UserList(UserIndex).Invent.Object(i)

            If .ObjIndex = ItemIndex Then
                If .Amount <= cant And .Equipped = 1 Then Call Desequipar(UserIndex, i)
                
                .Amount = .Amount - cant

                If .Amount <= 0 Then
                    cant = Abs(.Amount)
                    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
                    .Amount = 0
                    .ObjIndex = 0
                Else
                    cant = 0

                End If
                
                Call UpdateUserInv(False, UserIndex, i)
                
                If cant = 0 Then Exit Sub

            End If

        End With

    Next i

End Sub

Sub QuitarMateriales(ByVal UserIndex As Integer, _
                            ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Fecha: 20/08/2020
    'Descripcion: Quita la cantidad de materiales para construir
    '***************************************************
    Dim i As Byte
    With ObjData(ItemIndex)

        For i = 1 To MAXMATERIALES
            If .Materiales(i) > 0 Then Call QuitarObjetos(.Materiales(i), .CantMateriales(i), UserIndex)
        Next i

    End With

End Sub

Function TieneMateriales(ByVal UserIndex As Integer, _
                                   ByVal ItemIndex As Integer, _
                                   Optional ByVal ShowMsg As Boolean = False) As Boolean
    '***************************************************
    'Author: Lorwik
    'Fecha: 20/08/2020
    'Descripci�n: �Tiene materiales para construir?
    '***************************************************
    
    Dim i As Byte
    
    With ObjData(ItemIndex)

        For i = 1 To MAXMATERIALES
            If .Materiales(i) > 0 Then
                If Not TieneObjetos(.Materiales(i), .CantMateriales(i), UserIndex) Then
                    If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tienes suficiente materiales.", FontTypeNames.FONTTYPE_INFO)
                    TieneMateriales = False
                    Exit Function
    
                End If
    
            End If
        Next i
    
    End With

    TieneMateriales = True

End Function

Public Function PuedeConstruirItemHerrero(ByVal UserIndex As Integer, _
                               ByVal ItemIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 24/08/2009
    '24/08/2008: ZaMa - Validates if the player has the required skill
    '16/11/2009: ZaMa - Validates if the player has the required amount of materials, depending on the number of items to make
    '***************************************************
    PuedeConstruirItemHerrero = TieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(eSkill.herreria) >= ObjData(ItemIndex).SkHerreria

End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim i As Long

    For i = 1 To UBound(ArmasHerrero)

        If ArmasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    For i = 1 To UBound(ArmadurasHerrero)

        If ArmadurasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    PuedeConstruirHerreria = False

End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 30/05/2010
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '30/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************

    Dim TieneMateriales As Boolean

    Dim OtroUserIndex   As Integer

    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
            
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
            
                Call LimpiarComercioSeguro(UserIndex)

            End If

        End If
        
        'Sacamos energia
        'Chequeamos que tenga los puntos antes de sacarselos
        If .Stats.MinSta >= GASTO_ENERGIA Then
            .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(UserIndex) 'Paramos el macro
            Exit Sub

        End If

        
        Call QuitarMateriales(UserIndex, ItemIndex)
        ' AGREGAR FX
        
        'Mensajes de exito
        Select Case ObjData(ItemIndex).OBJType
            Case eOBJType.otWeapon
                Call WriteConsoleMsg(UserIndex, "Has construido el arma!.", FontTypeNames.FONTTYPE_INFO)
                    
            Case eOBJType.otEscudo
                Call WriteConsoleMsg(UserIndex, "Has construido el escudo!.", FontTypeNames.FONTTYPE_INFO)
                    
            Case eOBJType.otCasco
                Call WriteConsoleMsg(UserIndex, "Has construido el casco!.", FontTypeNames.FONTTYPE_INFO)
                    
            Case eOBJType.otArmadura
                Call WriteConsoleMsg(UserIndex, "Has construido la armadura!.", FontTypeNames.FONTTYPE_INFO)
        End Select
        
        Dim MiObj As obj
        
        MiObj.Amount = 1
        MiObj.ObjIndex = ItemIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then _
            Call TirarItemAlPiso(.Pos, MiObj)
        
        'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
        If ObjData(MiObj.ObjIndex).Log = 1 Then _
            Call LogDesarrollo(.name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
        
        Call SubirSkill(UserIndex, eSkill.herreria, True)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TRABAJO_HERRERO, .Pos.X, .Pos.Y))
        
        If Not criminal(UserIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
        
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

End Sub

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 28/05/2010
    '24/08/2008: ZaMa - Validates if the player has the required skill
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************
    On Error GoTo errHandler

    Dim TieneMateriales As Boolean

    Dim WeaponIndex     As Integer

    Dim OtroUserIndex   As Integer
    
    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)

            End If

        End If
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    
        If WeaponIndex <> SERRUCHO_CARPINTERO Then
            Call WriteConsoleMsg(UserIndex, "Debes tener equipado el serrucho para trabajar.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(UserIndex) 'Paramos el macro
            Exit Sub

        End If
    
        If .Stats.UserSkills(eSkill.Carpinteria) >= ObjData(ItemIndex).SkCarpinteria Then
           
            'Sacamos energia
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                Call DejardeTrabajar(UserIndex) 'Paramos el macro
                Exit Sub

            End If
            
            Call QuitarMateriales(UserIndex, ItemIndex)
            Call WriteConsoleMsg(UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
            
            Dim MiObj As obj

            MiObj.Amount = 1
            MiObj.ObjIndex = ItemIndex

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
            
            'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(.name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)

            End If
            
            Call SubirSkill(UserIndex, eSkill.Carpinteria, True)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TRABAJO_CARPINTERO, .Pos.X, .Pos.Y))
            
            If Not criminal(UserIndex) Then
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

                If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

            End If
            
            .Counters.Trabajando = .Counters.Trabajando + 1

        Else
            Call WriteConsoleMsg(UserIndex, "Aun no posees la habilidad suficiente para construir ese objeto. Necesitas al menos " & ObjData(ItemIndex).SkCarpinteria & " Skills.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub
errHandler:
    Call LogError("Error en CarpinteroConstruirItem. Error " & Err.Number & " : " & Err.description & ". UserIndex:" & UserIndex & ". ItemIndex:" & ItemIndex)

End Sub

Public Sub SastreConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 21/08/2020
    '***************************************************
    On Error GoTo errHandler

    Dim TieneMateriales As Boolean

    Dim WeaponIndex     As Integer

    Dim OtroUserIndex   As Integer
    
    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)

            End If

        End If
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    
        If WeaponIndex <> KIT_DE_COSTURA Then
            Call WriteConsoleMsg(UserIndex, "Debes tener equipado el kit de sastreria para trabajar.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(UserIndex) 'Paramos el macro
            Exit Sub

        End If
    
        If .Stats.UserSkills(eSkill.Sastreria) >= ObjData(ItemIndex).SkSastreria Then
           
            'Sacamos energia
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                Call DejardeTrabajar(UserIndex) 'Paramos el macro
                Exit Sub

            End If
            
            Call QuitarMateriales(UserIndex, ItemIndex)
            Call WriteConsoleMsg(UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
            
            Dim MiObj As obj

            MiObj.Amount = 1
            MiObj.ObjIndex = ItemIndex

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
            
            'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(.name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)

            End If
            
            Call SubirSkill(UserIndex, eSkill.Sastreria, True)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TRABAJO_CARPINTERO, .Pos.X, .Pos.Y))
            
            If Not criminal(UserIndex) Then
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

                If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

            End If
            
            .Counters.Trabajando = .Counters.Trabajando + 1

        Else
            Call WriteConsoleMsg(UserIndex, "Aun no posees la habilidad suficiente para construir ese objeto. Necesitas al menos " & ObjData(ItemIndex).SkSastreria & " Skills.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub
errHandler:
    Call LogError("Error en SastreConstruirItem. Error " & Err.Number & " : " & Err.description & ". UserIndex:" & UserIndex & ". ItemIndex:" & ItemIndex)

End Sub

Public Sub AlquimistaConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 21/08/2020
    '***************************************************
    On Error GoTo errHandler

    Dim TieneMateriales As Boolean

    Dim WeaponIndex     As Integer

    Dim OtroUserIndex   As Integer
    
    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)

            End If

        End If
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    
        If WeaponIndex <> OLLA_ALQUIMISTA Then
            Call WriteConsoleMsg(UserIndex, "Debes tener equipado la olla de alquimista para trabajar.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(UserIndex) 'Paramos el macro
            Exit Sub

        End If
    
        If .Stats.UserSkills(eSkill.Alquimia) >= ObjData(ItemIndex).SkAlquimia Then
           
            'Sacamos energia
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                Call DejardeTrabajar(UserIndex) 'Paramos el macro
                Exit Sub

            End If
            
            Call QuitarMateriales(UserIndex, ItemIndex)
            Call WriteConsoleMsg(UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
            
            Dim MiObj As obj

            MiObj.Amount = 1
            MiObj.ObjIndex = ItemIndex

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
            
            'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(.name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)

            End If
            
            Call SubirSkill(UserIndex, eSkill.Alquimia, True)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TRABAJO_CARPINTERO, .Pos.X, .Pos.Y))
            
            If Not criminal(UserIndex) Then
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

                If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

            End If
            
            .Counters.Trabajando = .Counters.Trabajando + 1

        Else
            Call WriteConsoleMsg(UserIndex, "Aun no posees la habilidad suficiente para construir ese objeto. Necesitas al menos " & ObjData(ItemIndex).SkAlquimia & " Skills.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub
errHandler:
    Call LogError("Error en AlquimistaConstruirItem. Error " & Err.Number & " : " & Err.description & ". UserIndex:" & UserIndex & ". ItemIndex:" & ItemIndex)

End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case Lingote

        Case iMinerales.HierroCrudo
            MineralesParaLingote = 14

        Case iMinerales.PlataCruda
            MineralesParaLingote = 20

        Case iMinerales.OroCrudo
            MineralesParaLingote = 35

        Case Else
            MineralesParaLingote = 10000

    End Select

End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
    '***************************************************
    '    Call LogTarea("Sub DoLingotes")
    Dim Slot           As Integer

    Dim obji           As Integer

    Dim CantidadItems  As Integer

    Dim TieneMinerales As Boolean

    Dim OtroUserIndex  As Integer
    
    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)

            End If

        End If
        
        CantidadItems = MaximoInt(1, CInt((.Stats.ELV - 4) / 5))

        Slot = .flags.TargetObjInvSlot
        obji = .Invent.Object(Slot).ObjIndex
        
        While CantidadItems > 0 And Not TieneMinerales

            If .Invent.Object(Slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
                TieneMinerales = True
            Else
                CantidadItems = CantidadItems - 1

            End If

        Wend
        
        If Not TieneMinerales Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - MineralesParaLingote(obji) * CantidadItems

        If .Invent.Object(Slot).Amount < 1 Then
            .Invent.Object(Slot).Amount = 0
            .Invent.Object(Slot).ObjIndex = 0

        End If
        
        Dim MiObj As obj

        MiObj.Amount = CantidadItems
        MiObj.ObjIndex = ObjData(.flags.TargetObjInvIndex).LingoteIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)

        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteConsoleMsg(UserIndex, "Has obtenido " & CantidadItems & " lingote" & IIf(CantidadItems = 1, "", "s") & "!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

End Sub

Function ModDomar(ByVal clase As eClass) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case clase

        Case eClass.Druid
            ModDomar = 6

        Case eClass.Hunter
            ModDomar = 6

        Case eClass.Cleric
            ModDomar = 7

        Case Else
            ModDomar = 10

    End Select

End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: 02/03/09
    '02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
    '***************************************************
    Dim j As Integer

    For j = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function

        End If

    Next j

End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    '***************************************************
    'Author: Nacho (Integer)
    'Last Modification: 01/05/2010
    '12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
    '02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
    '01/05/2010: ZaMa - Agrego bonificacion 11% para domar con flauta magica.
    '***************************************************

    On Error GoTo errHandler

    Dim puntosDomar      As Integer

    Dim puntosRequeridos As Integer

    Dim CanStay          As Boolean

    Dim petType          As Integer

    Dim NroPets          As Integer
    
    If Npclist(NpcIndex).MaestroUser = UserIndex Then
        Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    With UserList(UserIndex)

        If .NroMascotas < MAXMASCOTAS Then
            
            If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
                Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            If Not PuedeDomarMascota(UserIndex, NpcIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            puntosDomar = CInt(.Stats.UserAtributos(eAtributos.Carisma)) * CInt(.Stats.UserSkills(eSkill.Domar))
            
            ' 20% de bonificacion
            If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.8
            
                ' 11% de bonificacion
            ElseIf .Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
                puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.89
                
            Else
                puntosRequeridos = Npclist(NpcIndex).flags.Domable

            End If
            
            If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then

                Dim index As Integer

                .NroMascotas = .NroMascotas + 1
                index = FreeMascotaIndex(UserIndex)
                .MascotasIndex(index) = NpcIndex
                .MascotasType(index) = Npclist(NpcIndex).Numero
                
                Npclist(NpcIndex).MaestroUser = UserIndex
                
                Call FollowAmo(NpcIndex)
                Call ReSpawnNpc(Npclist(NpcIndex))
                
                Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
                
                ' Es zona segura?
                CanStay = (MapInfo(.Pos.Map).Pk = True)
                
                If Not CanStay Then
                    petType = Npclist(NpcIndex).Numero
                    NroPets = .NroMascotas
                    
                    Call QuitarNPC(NpcIndex)
                    
                    .MascotasType(index) = petType
                    .NroMascotas = NroPets
                    
                    Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. estas te esperaran afuera.", FontTypeNames.FONTTYPE_INFO)

                End If
                
                Call SubirSkill(UserIndex, eSkill.Domar, True)
        
            Else

                If Not .flags.UltimoMensaje = 5 Then
                    Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 5

                End If
                
                Call SubirSkill(UserIndex, eSkill.Domar, False)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "No puedes controlar mas criaturas.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, _
                                   ByVal NpcIndex As Integer) As Boolean

    '***************************************************
    'Author: ZaMa
    'This function checks how many NPCs of the same type have
    'been tamed by the user.
    'Returns True if that amount is less than two.
    '***************************************************
    Dim i           As Long

    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasType(i) = Npclist(NpcIndex).Numero Then
            numMascotas = numMascotas + 1

        End If

    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    'Makes an admin invisible o visible.
    '13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
    '***************************************************
    
    Dim tempData As String
    
    With UserList(UserIndex)

        If .flags.AdminInvisible = 0 Then

            ' Sacamos el mimetizmo
            If .flags.Mimetizado = 1 Then
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Counters.Mimetismo = 0
                .flags.Mimetizado = 0
                ' Se fue el efecto del mimetismo, puede ser atacado por npcs
                .flags.Ignorado = False

            End If
            
            'Guardamos el antiguo body y head
            .flags.OldBody = .Char.body
            .flags.OldHead = .Char.Head
            
            .flags.AdminInvisible = 1
            .flags.invisible = 1
            .flags.Oculto = 1
            
            ' Solo el admin sabe que se hace invi
            tempData = PrepareMessageSetInvisible(.Char.CharIndex, True)
            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(tempData)
            
            'Le mandamos el mensaje para que borre el personaje a los clientes que esten cerca
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
            
        Else
            .flags.AdminInvisible = 0
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            ' Solo el admin sabe que se hace visible
            tempData = PrepareMessageCharacterChange(.Char.body, .Char.Head, .Char.Heading, .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.loops, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(tempData)
            
            tempData = PrepareMessageSetInvisible(.Char.CharIndex, False)
            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(tempData)
             
            'Le mandamos el mensaje para crear el personaje a los clientes que esten cerca
            Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y, True)

        End If

    End With
    
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Suerte    As Byte

    Dim Exito     As Byte

    Dim obj       As obj

    Dim posMadera As WorldPos

    If Not LegalPos(Map, X, Y) Then Exit Sub

    With posMadera
        .Map = Map
        .X = X
        .Y = Y

    End With

    If MapData(Map, X, Y).ObjInfo.ObjIndex <> 58 Then
        Call WriteConsoleMsg(UserIndex, "Necesitas clickear sobre lena para hacer ramitas.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If MapData(Map, X, Y).ObjInfo.Amount < 3 Then
        Call WriteConsoleMsg(UserIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    Dim SupervivenciaSkill As Byte

    SupervivenciaSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia)

    If SupervivenciaSkill < 6 Then
        Suerte = 3
    ElseIf SupervivenciaSkill <= 34 Then
        Suerte = 2
    Else
        Suerte = 1

    End If

    Exito = RandomNumber(1, Suerte)

    If Exito = 1 Then
        obj.ObjIndex = FOGATA_APAG
        obj.Amount = MapData(Map, X, Y).ObjInfo.Amount \ 3
    
        Call WriteConsoleMsg(UserIndex, "Has hecho " & obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
        Call MakeObj(obj, Map, X, Y)
    
        'Seteamos la fogata como el nuevo TargetObj del user
        UserList(UserIndex).flags.TargetObj = FOGATA_APAG
    
        Call SubirSkill(UserIndex, eSkill.Supervivencia, True)
    Else

        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
            Call WriteConsoleMsg(UserIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 10

        End If

        '[/CDT]
    
        Call SubirSkill(UserIndex, eSkill.Supervivencia, False)

    End If

End Sub

Public Sub DoPescar(ByVal UserIndex As Integer, ByVal Red As Boolean)

    '***************************************************
    'Author: Unknown
    'Last Modification: 26/10/2018
    '26/10/2018: CHOTS - Multiplicador de oficios
    '***************************************************
    On Error GoTo errHandler

    Dim iSkill        As Integer

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim MAXITEMS      As Integer

    Dim CantidadItems As Integer

    With UserList(UserIndex)
    
        Call QuitarSta(UserIndex, ESFUERZOEXTRAER)

        iSkill = .Stats.UserSkills(eSkill.pesca)
        
        ' m = (60-11)/(1-10)
        ' y = mx - m*10 + 11
        
        Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

        If Suerte > 0 Then
            res = RandomNumber(1, Suerte)
            
            If res <= DificultadExtraer Then
            
                Dim MiObj As obj
                
                MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
                CantidadItems = RandomNumber(1, MAXITEMS)
                    
                CantidadItems = CantidadItems * OficioMultiplier
                
                MiObj.Amount = CantidadItems
                
                If Red Then
                    MiObj.ObjIndex = ListaPeces(RandomNumber(1, NUM_PECES))
                Else
                    MiObj.ObjIndex = Pescado
                End If
                
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)

                End If
                
                Call WriteConsoleMsg(UserIndex, "Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
                
                Call SubirSkill(UserIndex, eSkill.pesca, True)
            Else

                If Not .flags.UltimoMensaje = 6 Then
                    Call WriteConsoleMsg(UserIndex, "No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 6

                End If
                
                Call SubirSkill(UserIndex, eSkill.pesca, False)

            End If

        End If
        
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        
        'Sonido
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
        
        .Counters.Trabajando = .Counters.Trabajando + 1
    
    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en DoPescar Red: " & Red)

End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 05/04/2010
    'Last Modification By: ZaMa
    '24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
    '27/11/2009: ZaMa - Optimizacion de codigo.
    '18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
    '01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
    '05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
    '23/04/2010: ZaMa - No se puede robar mas sin energia.
    '23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
    '*************************************************

    On Error GoTo errHandler

    Dim OtroUserIndex As Integer

    If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub
    
    If UserList(VictimaIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    With UserList(LadrOnIndex)
    
        If .flags.Seguro Then
            If Not criminal(VictimaIndex) Then
                Call WriteConsoleMsg(LadrOnIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

        Else

            If .Faccion.ArmadaReal = 1 Then
                If Not criminal(VictimaIndex) Then
                    Call WriteConsoleMsg(LadrOnIndex, "Los miembros del ejercito real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub

                End If

            End If

        End If
        
        ' Caos robando a caos?
        If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de la legion oscura.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
        
        If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
        
        ' Tiene energia?
        If .Stats.MinSta < 15 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)

            End If
            
            Exit Sub

        End If
        
        ' Quito energia
        Call QuitarSta(LadrOnIndex, 15)
        
        Dim GuantesHurto As Boolean
    
        If .Invent.AnilloEqpObjIndex = GUANTE_HURTO Then GuantesHurto = True
        
        If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
            
            Dim Suerte     As Integer

            Dim res        As Integer

            Dim RobarSkill As Byte
            
            RobarSkill = .Stats.UserSkills(eSkill.Robar)
                
            If RobarSkill <= 10 Then
                Suerte = 35
            ElseIf RobarSkill <= 20 Then
                Suerte = 30
            ElseIf RobarSkill <= 30 Then
                Suerte = 28
            ElseIf RobarSkill <= 40 Then
                Suerte = 24
            ElseIf RobarSkill <= 50 Then
                Suerte = 22
            ElseIf RobarSkill <= 60 Then
                Suerte = 20
            ElseIf RobarSkill <= 70 Then
                Suerte = 18
            ElseIf RobarSkill <= 80 Then
                Suerte = 15
            ElseIf RobarSkill <= 90 Then
                Suerte = 10
            ElseIf RobarSkill < 100 Then
                Suerte = 7
            Else
                Suerte = 5

            End If
            
            res = RandomNumber(1, Suerte)
                
            If res < 3 Then 'Exito robo
                If UserList(VictimaIndex).flags.Comerciando Then
                    OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                        
                    If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                        Call WriteConsoleMsg(VictimaIndex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_TALK)
                        Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                        
                        Call LimpiarComercioSeguro(VictimaIndex)

                    End If

                End If
               
                If (RandomNumber(1, 50) < 25) And (.clase = eClass.Thief) Then
                    If TieneObjetosRobables(VictimaIndex) Then
                        Call RobarObjeto(LadrOnIndex, VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else 'Roba oro

                    If UserList(VictimaIndex).Stats.Gld > 0 Then

                        Dim n As Long
                        
                        If .clase = eClass.Thief Then

                            ' Si no tine puestos los guantes de hurto roba un 50% menos. Pablo (ToxicWaste)
                            If GuantesHurto Then
                                n = RandomNumber(.Stats.ELV * 50, .Stats.ELV * 100)
                            Else
                                n = RandomNumber(.Stats.ELV * 25, .Stats.ELV * 50)

                            End If

                        Else
                            n = RandomNumber(1, 100)

                        End If

                        If n > UserList(VictimaIndex).Stats.Gld Then n = UserList(VictimaIndex).Stats.Gld
                        UserList(VictimaIndex).Stats.Gld = UserList(VictimaIndex).Stats.Gld - n
                        
                        .Stats.Gld = .Stats.Gld + n

                        If .Stats.Gld > MAXORO Then .Stats.Gld = MAXORO
                        
                        Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & n & " monedas de oro a " & UserList(VictimaIndex).name, FontTypeNames.FONTTYPE_INFO)
                        Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
                        Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, True)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "" & .name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, False)

            End If
        
            If Not criminal(LadrOnIndex) Then
                If Not criminal(VictimaIndex) Then
                    Call VolverCriminal(LadrOnIndex)

                End If

            End If
            
            ' Se pudo haber convertido si robo a un ciuda
            If criminal(LadrOnIndex) Then
                .Reputacion.LadronesRep = .Reputacion.LadronesRep + vlLadron

                If .Reputacion.LadronesRep > MAXREP Then .Reputacion.LadronesRep = MAXREP

            End If

        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, _
                             ByVal Slot As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' Agregue los barcos
    ' Esta funcion determina que objetos son robables.
    ' 22/05/2010: Los items newbies ya no son robables.
    '***************************************************

    Dim OI As Integer

    OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

    ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And ObjData(OI).OBJType <> eOBJType.otBarcos And ObjData(OI).OBJType <> eOBJType.otMonturas And ObjData(OI).NoRobable = 1 And Not ItemNewbie(OI)

End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010
    '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
    '***************************************************

    Dim flag As Boolean

    Dim i    As Integer

    flag = False

    With UserList(VictimaIndex)

        If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
            i = 1

            Do While Not flag And i <= .CurrentInventorySlots

                'Hay objeto en este slot?
                If .Invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

                If Not flag Then i = i + 1
            Loop
        Else
            i = .CurrentInventorySlots

            Do While Not flag And i > 0

                'Hay objeto en este slot?
                If .Invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

                If Not flag Then i = i - 1
            Loop

        End If
    
        If flag Then

            Dim MiObj     As obj

            Dim Num       As Integer

            Dim ObjAmount As Integer
        
            ObjAmount = .Invent.Object(i).Amount
        
            'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
            Num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))
                                    
            MiObj.Amount = Num
            MiObj.ObjIndex = .Invent.Object(i).ObjIndex
        
            .Invent.Object(i).Amount = ObjAmount - Num
                    
            If .Invent.Object(i).Amount <= 0 Then
                Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)

            End If
                
            Call UpdateUserInv(False, VictimaIndex, CByte(i))
                    
            If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)

            End If
        
            If UserList(LadrOnIndex).clase = eClass.Thief Then
                Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name, FontTypeNames.FONTTYPE_INFO)

            End If

        Else
            Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningun objeto.", FontTypeNames.FONTTYPE_INFO)

        End If

        'If exiting, cancel de quien es robado
        Call CancelExit(VictimaIndex)
        
        'Si esta casteando, lo cancelamos
        Call CancelCast(VictimaIndex)

    End With

End Sub

Public Sub DoApunalar(ByVal UserIndex As Integer, _
                      ByVal VictimNpcIndex As Integer, _
                      ByVal VictimUserIndex As Integer, _
                      ByVal dano As Long)

    '***************************************************
    'Autor: Nacho (Integer) & Unknown (orginal version)
    'Last Modification: 04/17/08 - (NicoNZ)
    'Simplifique la cuenta que hacia para sacar la suerte
    'y arregle la cuenta que hacia para sacar el dano
    '***************************************************
    Dim Suerte As Integer

    Dim Skill  As Integer

    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apunalar)

    Select Case UserList(UserIndex).clase

        Case eClass.Assasin
            Suerte = Int(((0.00004 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
        Case eClass.Cleric, eClass.Paladin, eClass.Mercenario
            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
        Case eClass.Bard
            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
        Case Else
            Suerte = Int(0.0361 * Skill + 4.39)

    End Select

    If RandomNumber(0, 100) < Suerte Then
        If VictimUserIndex <> 0 Then
            If UserList(UserIndex).clase = eClass.Assasin Then
                dano = Round(dano * 1.4, 0)
            Else
                dano = Round(dano * 1.5, 0)

            End If
        
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, dano, DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(UserIndex, "Has apunalado a " & .name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, "Te ha apunalado " & UserList(UserIndex).name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)

            End With
        
        Else
            
            With Npclist(VictimNpcIndex)
                'Si el NPC es un Dummy no aplicamos el da�o
                'If Not .NPCtype = eNPCType.dummy Then
                '    .Stats.MinHp = .Stats.MinHp - Int(dano * 2)
                'End If
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, VictimNpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Int(dano * 2), DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(UserIndex, "Has apunalado la criatura por " & Int(dano * 2), FontTypeNames.FONTTYPE_FIGHT)
                Call CalcularDarExp(UserIndex, VictimNpcIndex, dano * 2)
            
            End With

        End If
    
        Call SubirSkill(UserIndex, eSkill.Apunalar, True)
    Else
        Call WriteConsoleMsg(UserIndex, "No has logrado apunalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
        Call SubirSkill(UserIndex, eSkill.Apunalar, False)

    End If

End Sub

Public Sub DoAcuchillar(ByVal UserIndex As Integer, _
                        ByVal VictimNpcIndex As Integer, _
                        ByVal VictimUserIndex As Integer, _
                        ByVal dano As Integer)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 12/01/2010
    '***************************************************

    If RandomNumber(1, 100) <= PROB_ACUCHILLAR Then
        dano = Int(dano * DANO_ACUCHILLAR)
        
        If VictimUserIndex <> 0 Then
        
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                Call WriteConsoleMsg(UserIndex, "Has acuchillado a " & .name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).name & " te ha acuchillado por " & dano, FontTypeNames.FONTTYPE_FIGHT)

            End With
            
        Else
            With Npclist(VictimNpcIndex)
            
                'Si el NPC es un Dummy no aplicamos el da�o
                'If Not .NPCtype = eNPCType.dummy Then
                '    .Stats.MinHp = .Stats.MinHp - dano
                'End If
                
                Call WriteConsoleMsg(UserIndex, "Has acuchillado a la criatura por " & dano, FontTypeNames.FONTTYPE_FIGHT)
                Call CalcularDarExp(UserIndex, VictimNpcIndex, dano)
            End With
        End If

    End If
    
End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, _
                          ByVal VictimNpcIndex As Integer, _
                          ByVal VictimUserIndex As Integer, _
                          ByVal dano As Long)

    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 28/01/2007
    '01/06/2010: ZaMa - Valido si tiene arma equipada antes de preguntar si es vikinga.
    '***************************************************
    Dim Suerte      As Integer

    Dim Skill       As Integer

    Dim WeaponIndex As Integer
    
    With UserList(UserIndex)

        ' Es bandido?
        If .clase <> eClass.Bandit Then Exit Sub
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
        
        ' Es una espada vikinga?
        If WeaponIndex <> ESPADA_VIKINGA Then Exit Sub
    
        Skill = .Stats.UserSkills(eSkill.Marciales)

    End With
    
    Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)
    
    If RandomNumber(1, 100) <= Suerte Then
    
        dano = Int(dano * 0.75)
        
        If VictimUserIndex <> 0 Then
            
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Int(dano * 2), DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(UserIndex, "Has golpeado criticamente a " & .name & " por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).name & " te ha golpeado criticamente por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)

            End With
            
        Else
            
            With Npclist(VictimNpcIndex)
                'Si el NPC es un Dummy no aplicamos el da�o
                'If Not .NPCtype = eNPCType.dummy Then
                '    .Stats.MinHp = .Stats.MinHp - dano
                'End If
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, VictimNpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Int(dano * 2), DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(UserIndex, "Has golpeado criticamente a la criatura por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)
                
                Call CalcularDarExp(UserIndex, VictimNpcIndex, dano)
            End With
            
           
            
        End If
        
    End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    'If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
    '    If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).Efecto = Trabajador Then _
    '        Cantidad = Porcentaje(Cantidad, 50)
    'End If

    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad

    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateSta(UserIndex)
    
    Exit Sub

errHandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.description)
    
End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex)
        .Counters.IdleCount = 0
        
        Dim Suerte       As Integer

        Dim res          As Integer

        Dim cant         As Integer

        Dim MeditarSkill As Byte
    
        'Barrin 3/10/03
        'Esperamos a que se termine de concentrar
        Dim TActual      As Long

        TActual = GetTickCount() And &H7FFFFFFF

        If TActual - .Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
            Exit Sub

        End If
        
        If .Counters.bPuedeMeditar = False Then
            .Counters.bPuedeMeditar = True

        End If
            
        If .Stats.MinMAN >= .Stats.MaxMAN Then
            Call WriteConsoleMsg(UserIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMeditateToggle(UserIndex)
            .flags.Meditando = False
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Exit Sub

        End If
        
        MeditarSkill = .Stats.UserSkills(eSkill.Meditar)
        
        If MeditarSkill <= 10 Then
            Suerte = 35
        ElseIf MeditarSkill <= 20 Then
            Suerte = 30
        ElseIf MeditarSkill <= 30 Then
            Suerte = 28
        ElseIf MeditarSkill <= 40 Then
            Suerte = 24
        ElseIf MeditarSkill <= 50 Then
            Suerte = 22
        ElseIf MeditarSkill <= 60 Then
            Suerte = 20
        ElseIf MeditarSkill <= 70 Then
            Suerte = 18
        ElseIf MeditarSkill <= 80 Then
            Suerte = 15
        ElseIf MeditarSkill <= 90 Then
            Suerte = 10
        ElseIf MeditarSkill < 100 Then
            Suerte = 7
        Else
            Suerte = 5

        End If

        res = RandomNumber(1, Suerte)
        
        If res = 1 Then
            
            cant = Porcentaje(.Stats.MaxMAN, PorcentajeRecuperoMana)
            
            '�Tiene anillo de sabiduria?
            'If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
            '    If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).Efecto = Sabiduria Then _
            '        cant = cant + 10
            'End If

            If cant <= 0 Then cant = 1
            .Stats.MinMAN = .Stats.MinMAN + cant

            If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
            
            'Renderizo el dano en render.
            Call WriteMessageCreateDamage(UserIndex, cant, DAMAGE_TRABAJO)
            
            Call WriteUpdateMana(UserIndex)
            Call SubirSkill(UserIndex, eSkill.Meditar, True)
        Else
            Call SubirSkill(UserIndex, eSkill.Meditar, False)

        End If

    End With

End Sub

Public Sub DoDesequipar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modif: 15/04/2010
    'Unequips either shield, weapon or helmet from target user.
    '***************************************************

    Dim Probabilidad   As Integer

    Dim Resultado      As Integer

    Dim WrestlingSkill As Byte

    Dim AlgoEquipado   As Boolean
    
    With UserList(UserIndex)

        ' Si no tiene guantes de hurto no desequipa.
        If .Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
        ' Si no esta solo con manos, no desequipa tampoco.
        If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
        WrestlingSkill = .Stats.UserSkills(eSkill.Marciales)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66

    End With
   
    With UserList(VictimIndex)

        ' Si tiene escudo, intenta desequiparlo
        If .Invent.EscudoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.WeaponEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.CascoEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
    
        If AlgoEquipado Then
            Call WriteConsoleMsg(UserIndex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "No has logrado desequipar ningun item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
    End With

End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modif: 03/03/2010
    'Implements the pick pocket skill of the Bandit :)
    '03/03/2010 - Pato: Solo se puede hurtar si no esta en trigger 6 :)
    '***************************************************
    Dim OtroUserIndex As Integer

    If TriggerZonaPelea(UserIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

    If UserList(UserIndex).clase <> eClass.Bandit Then Exit Sub

    'Esto es precario y feo, pero por ahora no se me ocurrio nada mejor.
    'Uso el slot de los anillos para "equipar" los guantes.
    'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
    If UserList(UserIndex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub

    Dim res As Integer

    res = RandomNumber(1, 100)

    If (res < 20) Then
        If TieneObjetosRobables(VictimaIndex) Then
    
            If UserList(VictimaIndex).flags.Comerciando Then
                OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                
                If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                    Call WriteConsoleMsg(VictimaIndex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_WARNING)
                    Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_WARNING)
                
                    Call LimpiarComercioSeguro(VictimaIndex)

                End If

            End If
                
            Call RobarObjeto(UserIndex, VictimaIndex)
            Call WriteConsoleMsg(VictimaIndex, "" & UserList(UserIndex).name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

        End If

    End If

End Sub

Public Sub DoHandInmo(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modif: 17/02/2007
    'Implements the special Skill of the Thief
    '***************************************************
    If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
    If UserList(UserIndex).clase <> eClass.Thief Then Exit Sub
    
    If UserList(UserIndex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
    Dim res As Integer

    res = RandomNumber(0, 100)

    If res < (UserList(UserIndex).Stats.UserSkills(eSkill.Marciales) / 4) Then
        UserList(VictimaIndex).flags.Paralizado = 1
        UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
        
        UserList(VictimaIndex).flags.ParalizedByIndex = UserIndex
        UserList(VictimaIndex).flags.ParalizedBy = UserList(UserIndex).name
        
        Call WriteParalizeOK(VictimaIndex)
        Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inmovil a tu oponente", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimaIndex, "El golpe te ha dejado inmovil!", FontTypeNames.FONTTYPE_FIGHT)

    End If

End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010 (ZaMa)
    '02/04/2010: ZaMa - Nueva formula para desarmar.
    '***************************************************

    Dim Probabilidad   As Integer

    Dim Resultado      As Integer

    Dim WrestlingSkill As Byte
    
    With UserList(UserIndex)
        WrestlingSkill = .Stats.UserSkills(eSkill.Marciales)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
        
        Resultado = RandomNumber(1, 100)
        
        If Resultado <= Probabilidad Then
            Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
            Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

            If UserList(VictimIndex).Stats.ELV < 20 Then
                Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If

    End With
    
End Sub

Public Function MaxItemsConstruibles(ByVal UserIndex As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '***************************************************
    
    With UserList(UserIndex)

    MaxItemsConstruibles = MaximoInt(1, CInt((.Stats.ELV - 2) * 0.2))

    End With

End Function

Public Function MaxItemsExtraibles(ByVal UserLevel As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/05/2010
    '***************************************************
    MaxItemsExtraibles = MaximoInt(1, CInt((UserLevel - 2) * 0.2)) + 1

End Function

Public Sub ImitateNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    'Copies body, head and desc from previously clicked npc.
    '***************************************************
    
    With UserList(UserIndex)
        
        ' Copy desc
        .DescRM = Npclist(NpcIndex).name
        
        ' Remove Anims (Npcs don't use equipment anims yet)
        .Char.CascoAnim = NingunCasco
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        
        ' If admin is invisible the store it in old char
        If .flags.AdminInvisible = 1 Or .flags.invisible = 1 Or .flags.Oculto = 1 Then
            
            .flags.OldBody = Npclist(NpcIndex).Char.body
            .flags.OldHead = Npclist(NpcIndex).Char.Head
        Else
            .Char.body = Npclist(NpcIndex).Char.body
            .Char.Head = Npclist(NpcIndex).Char.Head
            
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

        End If
    
    End With
    
End Sub

Public Sub DoEquita(ByVal UserIndex As Integer, _
                    ByRef Montura As ObjData, _
                    ByVal Slot As Integer)
    '***************************************************
    'Author: Recox
    'Last Modification: 06/04/2020
    'Podemos usar monturas ahora
    '06/04/2020: FrankoH298 - Ahora hay un timer para poder montarte
    '***************************************************

    With UserList(UserIndex)
    
        If UserList(UserIndex).Stats.UserSkills(Equitacion) < Montura.MinSkill Then
            Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitaci�n.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        '�Esta intentando usar una montura de tipo dungeon fuera de un dungeon?
        'If MapInfo(.Pos.Map).Zona <> "DUNGEON" And Montura.MontTipo = 1 Then
        '    Call WriteConsoleMsg(UserIndex, "No puedes utilizar esta montura fuera de un dungeon.", FontTypeNames.FONTTYPE_INFO)
        '    Exit Sub
        'End If

        '�Esta en un dungeon y la montura no es de tipo dungeon?
        'If MapInfo(.Pos.Map).Zona = "DUNGEON" And Montura.MontTipo <> 1 Then
        '    Call WriteConsoleMsg(UserIndex, "No puedes utilizar esta montura en dungeon.", FontTypeNames.FONTTYPE_INFO)
        '    Exit Sub
        'End If
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes utilizar la montura mientras estas muerto !!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.Navegando = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes utilizar la montura mientras navegas !!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.BAJOTECHO Or MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.CASA Then
            'TODO: SACAR ESTA VALIDACION DE ACA, Y HACER UN legalpos HAY TECHO en el cliente
            If .flags.Equitando = 0 Then Exit Sub

            Call WriteConsoleMsg(UserIndex, "No puedes utilizar la montura bajo techo!", FontTypeNames.FONTTYPE_INFO)
        End If

        ' If .flags.Metamorfosis = 1 Then 'Metamorfosis
        '     Call WriteConsoleMsg(UserIndex, "No puedes montar mientras estas metamorfoseado.", FontTypeNames.FONTTYPE_INFO)
        '     Exit Sub
        ' End If

        ' No estaba equitando
        If .flags.Equitando = 0 Then

            If .Counters.MonturaCounter <= 0 Then
                .Invent.MonturaObjIndex = .Invent.Object(Slot).ObjIndex
                .Invent.MonturaEqpSlot = Slot
    
                Call ToggleMonturaBody(UserIndex)
                Call SetVisibleStateForUserAfterNavigateOrEquitate(UserIndex)
    
                '  Comienza a equitar
                .flags.Equitando = 1
                .flags.Velocidad = ObjData(.Invent.MonturaObjIndex).Speed
                Call WriteSetSpeed(UserIndex)
                
                Call WriteEquitandoToggle(UserIndex)

                'Mostramos solo el casco de los items equipados por que los demas items quedan mal en el render, solo es un tema visual (Recox)
                Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
            Else
                Call WriteConsoleMsg(UserIndex, "Debe esperar " & .Counters.MonturaCounter & " segundos para volver a usar tu montura", FontTypeNames.FONTTYPE_INFO)
            End If
            
        ' Estaba equitando
        Else
            Call UnmountMontura(UserIndex)
            Call WriteEquitandoToggle(UserIndex)

        End If


    End With

End Sub

Public Sub UnmountMontura(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .Invent.MonturaObjIndex = 0
        .Invent.MonturaEqpSlot = 0

        .Char.Head = .OrigChar.Head

        ' Seteamos el equipo que tiene y lo mostramos en el render.
        Call SetEquipmentOnCharAfterNavigateOrEquitate(UserIndex)
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
  
        ' Termina de equitar
        .flags.Equitando = 0
        .flags.Velocidad = SPEED_NORMAL
        Call WriteSetSpeed(UserIndex)
        
        .Counters.MonturaCounter = 3

    End With
End Sub

Private Sub SetVisibleStateForUserAfterNavigateOrEquitate(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        ' Pierde el ocultar
        If .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .Counters.Ocultando = 0
            Call SetInvisible(UserIndex, .Char.CharIndex, False)
            Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
        End If

        ' Siempre se ve la montura (Nunca esta invisible), pero solo para el cliente.
        If .flags.invisible = 1 Then
            Call SetInvisible(UserIndex, .Char.CharIndex, False)
        End If

    End With

End Sub

Private Sub SetEquipmentOnCharAfterNavigateOrEquitate(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        If .Invent.ArmourEqpObjIndex > 0 Then
            .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)

        End If
        
        If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

        If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Invent.WeaponEqpObjIndex)

        If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
        
    End With


End Sub

Public Sub DoExtraer(ByVal UserIndex As Integer, ByVal Profesion As Integer)

    '***************************************************
    'Autor: Lorwik
    'Fecha: 19/08/2020
    'Descripci�n: Extrae recursos de forma pasiva
    '***************************************************
    
    On Error GoTo errHandler

    Dim Suerte        As Integer
    Dim res           As Integer
    Dim MAXITEMS      As Integer
    Dim CantidadItems As Integer
    Dim MiObj As obj
    

    With UserList(UserIndex)

        If .flags.TargetObj = 0 Then Exit Sub

        '�La herramienta es de la misma categoria o superior?
        If ObjData(.flags.TargetObj).Recurso.Categoria > ObjData(.Invent.WeaponEqpObjIndex).Herramienta.Categoria Then
            Call WriteConsoleMsg(UserIndex, "El recurso que intentas extraer es demasiado duro para esa herramienta.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(UserIndex) 'Paramos el macro
            Exit Sub
        End If

        Call QuitarSta(UserIndex, ESFUERZOEXTRAER)

        Dim Skill As Integer

        Skill = .Stats.UserSkills(Profesion)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)

        If res <= DificultadExtraer Then
        
            MiObj.ObjIndex = ObjData(.flags.TargetObj).RecursoIndex
        
            MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
            
            CantidadItems = RandomNumber(1, MAXITEMS)

            CantidadItems = CantidadItems * OficioMultiplier

            MiObj.Amount = CantidadItems
       
            If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
        
            Call WriteConsoleMsg(UserIndex, "Has extraido algunos materiales!", FontTypeNames.FONTTYPE_INFO)
            
            'Renderizo el dano en render.
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, MiObj.Amount, DAMAGE_TRABAJO))
            Call WriteMessageCreateDamage(UserIndex, MiObj.Amount, DAMAGE_TRABAJO)
            
            Call SubirSkill(UserIndex, Profesion, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 9 Then
                Call WriteConsoleMsg(UserIndex, "No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 9

            End If

            '[/CDT]
            Call SubirSkill(UserIndex, Profesion, False)

        End If
    
        If Not criminal(UserIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1
        
        'Play sound!
        If Profesion = eSkill.Mineria Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_MINERO, .Pos.X, .Pos.Y))
            
        ElseIf Profesion = eSkill.Talar Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
            
        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en Sub DoExtraer")

End Sub

' <<<<<< ------ INSTRUCTORES ------ >>>>>>

Public Sub AccionInstructor(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    '***************************************************
    'Autor: Lorwik
    'Fecha: 19/08/2020
    'Descripci�n: �El usuario quiere aprender u olvidar una profesion?
    '***************************************************
    
    Dim SlotLibre As Boolean
    Dim i As Integer
    
    With UserList(UserIndex)
    
        '�Instruye una profesion valida?
        If Npclist(NpcIndex).Instruye <= 0 Then
            Call WriteConsoleMsg(UserIndex, "El instructor esta enfermo y no puede instruirte en la profesi�n.", FontTypeNames.FONTTYPE_INFO) 'Excusa para el user xD
            Exit Sub
        End If
        
        '�Desea aprender?
        If ConoceProfesion(UserIndex, Npclist(NpcIndex).Instruye) < 0 Then
        
            '�Tiene slot libre para aprender una profesion?
            For i = 0 To 1
                If .Profesion(i).Profesion = 0 Then SlotLibre = True
            Next i
            
            If SlotLibre = False Then
                Call WriteConsoleMsg(UserIndex, "Ya conoces 2 profesiones, si quieres aprender esta profesion debes olvidar alguna de las ya conocidas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            Call WriteConfirmarInstruccion(UserIndex, "�Seguro que quieres instruirte en " & SkillsNames(Npclist(NpcIndex).Instruye) & "?, el precio son " & PRECIOINSTRUCCION & " monedas de oro.")
            .flags.ProfInstruyendo = Npclist(NpcIndex).Instruye
            .flags.Instruyendo = 1 '1: Aprender
           
        Else 'Entonces quiere olvidar
        
            Call WriteConfirmarInstruccion(UserIndex, "�Seguro que quieres olvidar la profesion de " & SkillsNames(Npclist(NpcIndex).Instruye) & "?, perderas todos los skills y TODAS las RECETAS adquiridas en dicha profesion.")
            .flags.ProfInstruyendo = Npclist(NpcIndex).Instruye
            .flags.Instruyendo = 2 '2: Olvidar
            
        End If
        
    End With
    
End Sub

Public Sub AccionProfesion(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 19/08/2020
'Descripci�n: Aprende u olvida una profesion
'1: Aprender
'2: Olvidar
'***************************************************
    Dim i As Byte
    Dim Slot As Byte
    
    With UserList(UserIndex)
    
        '�Esta instruyendose?
        If .flags.Instruyendo = 1 Then
        
            'La instruccion cuesta 5K
            If UserList(UserIndex).Stats.Gld < PRECIOINSTRUCCION Then
                Call WriteConsoleMsg(UserIndex, "No suficiente dinero para pagar al instructor. Necesitas " & PRECIOINSTRUCCION & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Buscamos un hueco libre
            For i = 0 To 1
                If .Profesion(i).Profesion = 0 Then Slot = i
            Next i
            
            .Profesion(Slot).Profesion = .flags.ProfInstruyendo
            
            'Si es una profesion de crafting le damos una receta inicial:
            Select Case .flags.ProfInstruyendo
            
                Case eSkill.herreria
                    .Profesion(Slot).Recetas(1) = 15 'Daga
                    
                Case eSkill.Carpinteria
                    .Profesion(Slot).Recetas(1) = 163 'Cuchara
                    
                Case eSkill.Alquimia
                    .Profesion(Slot).Recetas(1) = 166 'Pocion Violeta
                    
                Case eSkill.Sastreria
                    .Profesion(Slot).Recetas(1) = 641 'Ropa comun
            
            End Select
            
            .Stats.UserSkills(.flags.ProfInstruyendo) = .Stats.UserSkills(.flags.ProfInstruyendo) + 1
            Call CheckEluSkill(UserIndex, .flags.ProfInstruyendo, True)
            
            'Restamos el oro
            .Stats.Gld = .Stats.Gld - PRECIOINSTRUCCION
            Call WriteUpdateGold(UserIndex)
            
            Call WriteConsoleMsg(UserIndex, "�Bienvenido al gremio de " & SkillsNames(.flags.ProfInstruyendo) & "! queda en tus manos adquirir mas destreza en la profesion", FontTypeNames.FONTTYPE_INFO)
            
            'Reseteamos los flags
            .flags.ProfInstruyendo = 0
            .flags.Instruyendo = 0
    
        ElseIf .flags.Instruyendo = 2 Then '�Esta olvidando?
        
            'Buscamos y olvidamos la profesion
            Slot = ConoceProfesion(UserIndex, .flags.ProfInstruyendo)
            
            .Profesion(Slot).Profesion = 0
            
            'Eliminamos todas las recetas
            For i = 1 To MAXUSERRECETAS
                .Profesion(Slot).Recetas(i) = 0
            Next i
        
            Call WriteConsoleMsg(UserIndex, "Es una lastima que hayas decidido abandonar el gremio de " & SkillsNames(.flags.ProfInstruyendo) & ".", FontTypeNames.FONTTYPE_INFO)
        
            'Eliminamos los skills
            .Stats.UserSkills(.flags.ProfInstruyendo) = 0
            
            'Reseteamos los flags
            .flags.ProfInstruyendo = 0
            .flags.Instruyendo = 0
            
        End If
    
    End With
    
End Sub

Public Function ConoceProfesion(ByVal UserIndex As Integer, ByVal Profesion As Byte) As Integer
'***************************************************
'Autor: Lorwik
'Fecha: 19/08/2020
'Descripci�n: Obtiene el Slot de la profesion, si no la consigue es que no la tiene
'***************************************************

    Dim i As Byte
    
    For i = 0 To 1
        If UserList(UserIndex).Profesion(i).Profesion = Profesion Then
            ConoceProfesion = i
            Exit Function
        End If
    Next i

    ConoceProfesion = -1

End Function

Sub AgregarReceta(ByVal UserIndex As Integer, ByVal Slot As Integer)
    '***************************************************
    'Autor: Lorwik
    'Fecha: 21/08/2020
    'Descripci�n: Agregamos una receta, patron o lo que sea de una profesion al conocimiento del user
    '***************************************************

    Dim rIndex          As Integer
    Dim j               As Integer
    Dim SlotProfesion   As Integer

    With UserList(UserIndex)
    
        SlotProfesion = ConoceProfesion(UserIndex, ObjData(.Invent.Object(Slot).ObjIndex).Profesion)
    
        '�Tiene la profesion de la receta?
        If SlotProfesion < 0 Then
            Call WriteConsoleMsg(UserIndex, "Intentas leer el pergamino, pero todo te resulta desconocido. No conoces esa profesi�n.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        rIndex = ObjData(.Invent.Object(Slot).ObjIndex).RecetaIndex
    
        If TieneReceta(rIndex, UserIndex, SlotProfesion) = False Then

            'Buscamos un slot vacio
            For j = 1 To MAXUSERRECETAS

                If .Profesion(SlotProfesion).Recetas(j) = 0 Then Exit For
            Next j
            
            If .Profesion(SlotProfesion).Recetas(j) <> 0 Then
                Call WriteConsoleMsg(UserIndex, "No tienes espacio para mas recetas.", FontTypeNames.FONTTYPE_INFO)
                
            Else
                .Profesion(SlotProfesion).Recetas(j) = rIndex

                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Ya tienes esa receta.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

Function TieneReceta(ByVal i As Integer, ByVal UserIndex As Integer, ByVal SlotProfesion As Byte) As Boolean
    '***************************************************
    'Autor: Lorwik
    'Fecha: 21/08/2020
    'Descripcion: Busca una receta entre las conocidas
    '***************************************************

    On Error GoTo errHandler
    
    Dim j As Integer

    With UserList(UserIndex)
    
        For j = 1 To MAXUSERRECETAS

            If .Profesion(SlotProfesion).Recetas(j) = i Then
                TieneReceta = True
                Exit Function
    
            End If
    
        Next
        
    End With
    
    TieneReceta = False
    Exit Function
errHandler:

End Function


