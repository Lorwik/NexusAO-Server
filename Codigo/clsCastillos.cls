VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCastillos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Nombre_Castillo      As String

Private Mapa_Castillo        As Byte

Private Rey_Castillo         As Integer

Private Dueno                As String

Private SpawnX               As Byte

Private SpawnY               As Byte

Private CastilloID           As Byte

Private Const GLD_RECOMPENSA As Long = 100000

Public Sub Inicializar(ByVal ID As Byte, _
                       ByVal Nombre As String, _
                       ByVal Mapa As Byte, _
                       ByVal Rey As Integer, _
                       ByVal X As Byte, _
                       ByVal Y As Byte)

'*********************************
'Autor: Lorwik
'Fecha: 28/05/2023
'*********************************

    CastilloID = ID
    Nombre_Castillo = Nombre
    Mapa_Castillo = Mapa
    Rey_Castillo = Rey
    SpawnX = X
    SpawnY = Y
    
    Call SpawnRey
    
End Sub

Public Property Let setConquistador(ByVal Nombre As String)
    Dueno = Nombre

End Property

Public Property Get getConquistador() As String
    getConquistador = Dueno

End Property

Public Property Get getMapa() As Byte
    getMapa = Mapa_Castillo

End Property

Public Property Get getRey() As Integer
    getRey = Rey_Castillo

End Property

Public Property Get getNombreCastillo() As String
    getNombreCastillo = Nombre_Castillo

End Property

Private Sub SpawnRey()
'*********************************
'Autor: Lorwik
'Fecha: 28/05/2023
'Descripción: Spawnea el rey en el castillo o fortaleza
'*********************************

    Dim ReyPos As WorldPos
    
    ReyPos.Map = Mapa_Castillo
    ReyPos.X = SpawnX
    ReyPos.Y = SpawnY
    
    Call SpawnNpc(Rey_Castillo, ReyPos, False, False)
    
End Sub

Public Sub MuereRey(ByVal UserIndex As Integer)
'*********************************
'Autor: Lorwik
'Fecha: 28/05/2023
'Descripción: Acciones cuando el rey muere
'*********************************

    Dueno = modGuilds.GuildName(UserList(UserIndex).GuildIndex)

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Nombre_Castillo & " ha sido tomado por el clan " & Dueno, FontTypeNames.FONTTYPE_SERVER))
    
    'Guardamos el nuevo dueño
    Call WriteVar(DatPath & "\Castillos.dat", "CASTLE" & CastilloID, "Clan", Dueno)
    
    'El Rey reaparece
    Call SpawnRey

End Sub

Private Sub darRecompensa(ByVal UserIndex As Integer)
'*********************************
'Autor: Lorwik
'Fecha: 28/05/2023
'Descripción: Se le recompensa al usuario por mantener el castillo o fortaleza bajo su dominio
'*********************************

    With UserList(UserIndex)
    
        .Stats.Gld = .Stats.Gld + GLD_RECOMPENSA
        
        Call WriteUpdateGold(UserIndex)
        
        Call WriteConsoleMsg(UserIndex, "Has recibido una recompensa por mantener el castillo de " & GLD_RECOMPENSA & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
    
    End With

End Sub

Public Sub RecompensarGuild()
'*********************************
'Autor: Lorwik
'Fecha: 28/05/2023
'Descripción: Busca a todos los miembros del clan dominante para recompensarlos.
'*********************************

    Dim LoopC As Integer

    For LoopC = 1 To LastUser
        
        '¿Tiene clan?
        If UserList(LoopC).GuildIndex > 0 Then
            '¿Su clan tomo el castillo?
            If guilds(UserList(LoopC).GuildIndex).GuildName = Dueno Then Call darRecompensa(LoopC)
        End If
        
    Next LoopC

End Sub
