Attribute VB_Name = "modELO"
Option Explicit

Private Type Rank

    nombre As String
    ELO As Double
    Posicion As Byte

End Type

Public Ranked(5) As Rank

Public Function CalcularELO(ByVal UserA As Integer, _
                            ByVal UserB As Integer, _
                            ByVal Gana As Boolean) As Double

    Dim ELOUserA      As Double

    Dim ELOUserB      As Double

    Dim ELODiferencia As Double

    Dim FactorK       As Byte

    Dim Elevado       As Double

    Dim Porcentaje    As Double
    
    ELOUserA = UserList(UserA).Stats.ELO
    ELOUserB = UserList(UserB).Stats.ELO
    
    FactorK = 32
    
    ELODiferencia = ELOUserB - ELOUserA
    
    Elevado = ELODiferencia / 400
    Porcentaje = 1 / (1 + 10 ^ Elevado)

    If Gana = True Then
        'Gana
        CalcularELO = (FactorK * (1 - Porcentaje))
    ElseIf Gana = False Then
        'Pierde
        CalcularELO = (FactorK * (0 - Porcentaje))

    End If
    
End Function

Public Sub CargarRank()

    On Error GoTo errHandler

    Dim Leer As clsIniManager

    Set Leer = New clsIniManager

    Dim i As Byte
        
    Call Leer.Initialize(DatPath & "\Ranking.dat")
        
    For i = 1 To 5
        Ranked(i).nombre = Leer.GetValue("Posicion" & i, "Nombre")
        Ranked(i).ELO = Leer.GetValue("Posicion" & i, "ELO")
        Ranked(i).Posicion = i
    Next i
        
    Set Leer = Nothing
    Exit Sub

errHandler:
    MsgBox "Error cargando Ranking.dat " & Err.Number & ": " & Err.description

End Sub

Public Sub GuardarRank()

    Dim i    As Byte

    Dim File As String
    
    File = DatPath & "\Ranking.dat"
        
    For i = 1 To 5
        Call WriteVar(File, "Posicion" & i, "Nombre", Ranked(i).nombre)
        Call WriteVar(File, "Posicion" & i, "ELO", Ranked(i).ELO)
    Next i

End Sub

Public Sub ActualizarRank(ByVal UserIndex As Integer)

    Dim i            As Byte

    Dim ELOIndex     As Double

    Dim NameIndex    As String

    Dim ViejoELO     As Double

    Dim ViejoNombre  As String

    Dim ViejaPos     As Byte

    Dim UserAgregado As Boolean
    
    ELOIndex = UserList(UserIndex).Stats.ELO
    NameIndex = UserList(UserIndex).name
    
    For i = 1 To 5

        If (i = 5) And (UserAgregado = True) Then Exit Sub
        If UserAgregado Then
            If i + 1 < 5 Then
                Ranked(i + 1).nombre = NameIndex
                Ranked(i + 1).ELO = ELOIndex
                Ranked(i + 1).Posicion = i

            End If

        End If

        If Ranked(i).ELO <= ELOIndex Then
            If i + 1 < 5 Then
                Ranked(i + 1).nombre = Ranked(i).nombre
                Ranked(i + 1).ELO = Ranked(i).ELO
                Ranked(i + 1).Posicion = i + 1
                
                'Insertamos al usuario en su nueva posicion
                Ranked(i).nombre = NameIndex
                Ranked(i).ELO = ELOIndex
                Ranked(i).Posicion = i
                UserAgregado = True
                Call GuardarRank

            End If

        End If

    Next i

End Sub

