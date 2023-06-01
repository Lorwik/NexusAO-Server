Attribute VB_Name = "modCastillos"
Option Explicit

Public TimerRecompensaCastillos As Long

Public Sub SetearContador()
'*****************************************
'Autor: Lorwik
'Fecha: 28/05/2026
'Descripción: Setea el contador de recompensas de los castillos
'*****************************************

    TimerRecompensaCastillos = 3060 '1 hora

End Sub

Public Sub RepartirRecompensas()
'*****************************************
'Autor: Lorwik
'Fecha: 28/05/2026
'Descripción: Reparte las recompensas
'*****************************************

    Dim i As Byte

    For i = 1 To CastleCount
        Call Castillo(i).RecompensarGuild
    Next i
End Sub

Public Function obtenerCastilloporMapa(ByVal Mapa As Integer) As Byte
'*****************************************
'Autor: Lorwik
'Fecha: 28/05/2026
'Descripción: Obtiene la Id del castillo segun el mapa que se le mando
'*****************************************

    Dim i As Byte
    
    For i = 1 To CastleCount
    
        If Mapa = Castillo(i).getMapa Then
            obtenerCastilloporMapa = i
            Exit Function
        End If
    
    Next i

End Function
