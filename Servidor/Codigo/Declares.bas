Attribute VB_Name = "Declaraciones"

Option Explicit

Public TD_A As Byte
Public TD_B As Byte
Public HayTD As Boolean
Public EquipoTemp As Byte
Public SlotsTD As Byte
Public MaxSlotsTD As Byte

'----------------------------------------------------------------
'Aca guardamos todas las posiciones de mierda esas
'----------------------------------------------------------------

Public PosUserPareja1 As WorldPos
Public PosUserPareja2 As WorldPos
Public PosUserPareja3 As WorldPos
Public PosUserPareja4 As WorldPos

Public PosUserSum As WorldPos
Public PosUserConsulta As WorldPos

Public PosUserDuelo1 As WorldPos
Public PosUserDuelo2 As WorldPos

Public PosUserReto1 As WorldPos
Public PosUserReto2 As WorldPos

'----------------------------------------------------------------
'/Aca guardamos todas las posiciones de mierda esas
'----------------------------------------------------------------

Public AutoTorneo As Integer
Public AutoMensaje As Byte
Public ReyON As Byte
Public MurioDragon As Byte
Public GuardiasRey As Byte

Public ObjSlot1 As Byte
Public ObjSlot2 As Byte

Public PublicKey As Integer
Public PrivateKey As Integer

Public CvcFunciona As Boolean
Public UsuariosEnTorneo As Integer

Public PasoHD As Boolean
Public HDSerialIndex As String

Public MEMBERSFILE                 As String       'decente la capa de arriba se entera donde estan
Public Const mapainvo = 93
Public Const maparey = 106
Public Const mapainvoX1 = 29
Public Const mapainvoY1 = 26
Public Const mapainvoX2 = 26
Public Const mapainvoY2 = 29
Public Const mapainvoX3 = 29
Public Const mapainvoY3 = 32
Public Const mapainvoX4 = 32
Public Const mapainvoY4 = 29

'[Loopzer]
Public Lac_Camina As Long
Public Lac_Pociones As Long
Public Lac_Pegar As Long
Public Lac_Lanzar As Long
Public Lac_Usar As Long
Public Lac_Tirar As Long
 
Public Type TLac
    LCaminar As New Cls_InterGTC
    LPociones As New Cls_InterGTC
    LPegar As New Cls_InterGTC
    LUsar As New Cls_InterGTC
    LTirar As New Cls_InterGTC
    LLanzar As New Cls_InterGTC
End Type
'[/Loopzer]


Public LimpiezaTimerMinutos As Byte
Public Const TimerCleanWorld As Byte = 60  'No superar los 255 minutos, esto es cada cuanto quieren que se realice un limpiado de mundo.

Public CuentaTorneo As Integer
Public contador As Long

Public CuentaRegresiva As Long
Public CuentaArena As Long

Public CastilloNorte As String
Public CastilloSur As String
Public CastilloEste As String
Public CastilloOeste As String 'By Nait

Public NombreUsuariosMatados As String
Public UsuariosMatadosCantidad As Integer
Public NombrePuntos As String
Public PuntosDeTorneo As Long
Public NombreRepu As String
Public Repu As Long
Public NombreTrofeos As String
Public TrofeosDeOro As Integer
Public Oro As Long
Public NombreRetos As String
Public RetosGaGanados As Long
Public NombreDuelos As String
Public DuelosGaGanados As Long

Public HayPareja As Boolean

Public Consulta As Boolean
Public HayConsulta As Boolean

Public Arena1 As Boolean
Public Arena2 As Boolean
Public Arena3 As Boolean
Public Arena4 As Boolean

Type tEstadisticasDiarias
    Segundos As Double
    MaxUsuarios As Integer
    Promedio As Integer
End Type
    
Public DayStats As tEstadisticasDiarias

Public aDos As New clsAntiDos

Public aClon As New clsAntiMassClon
Public TrashCollector As New Collection

Public EHWACHIN As Byte

Public Const MAXSPAWNATTEMPS = 60
Public Const MAXUSERMATADOS = 9000000
Public Const LoopAdEternum = 999
Public Const FXSANGRE = 14
Public Const FXAPUÑALAR = 54


Public Const iFragataFantasmal = 87

Public Enum iMinerales
    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388
End Enum

Public Enum PlayerType
    User = 0
    VIP = 1
    SemiDios = 2
    Dios = 3
    Admin = 4
End Enum

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Const LimiteNewbie As Byte = 12

'Barrin 3/10/03
Public Const TIEMPO_INICIOMEDITAR As Byte = 1

Public Const NingunEscudo As Integer = 2
Public Const NingunCasco As Integer = 2
Public Const NingunArma As Integer = 2

Public Const EspadaMataDragones As Integer = 1053
Public Const EspadaMataDragonesIndex As Integer = 402
Public Const LAUDMAGICO As Integer = 696

Public Const MAXMASCOTASENTRENADOR As Byte = 7

Public Enum FXIDs
    FXWARP = 1
    FXMEDITARCHICO = 4
    FXMEDITARMEDIANO = 5
    FXMEDITARGRANDE = 6
    FXMEDITARXGRANDE = 43
    FXMEDITARCIUDA = 44
    FXMEDITARCRIMI = 42
    FXMEDITARVIP = 44
    FXMEDITARVIPW = 45
    FXMEDITARTRANSFO = 16
End Enum

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

''
' constantes para el trigger 6
'
' @see eTrigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum eTrigger6
    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3
End Enum

'TODO : Reemplazar por un enum
Public Const Bosque = "BOSQUE"
Public Const Nieve = "NIEVE"
Public Const Desierto = "DESIERTO"
Public Const Ciudad = "CIUDAD"
Public Const Campo = "CAMPO"
Public Const Dungeon = "DUNGEON"

' <<<<<< Targets >>>>>>
Public Enum TargetType
    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4
    uOnlyUsuario = 5
End Enum

' <<<<<< Acciona sobre >>>>>>
Public Enum TipoHechizo
    uPropiedades = 1
    uEstado = 2
    uMaterializa = 3    'Nose usa
    uInvocacion = 4
End Enum

Public Const DRAGON As Integer = 6

Public Const MAX_MENSAJES_FORO As Byte = 35

Public Const MAXUSERHECHIZOS As Byte = 11


' TODO: Y ESTO ? LO CONOCE GD ?
Public Const EsfuerzoTalarGeneral As Byte = 4
Public Const EsfuerzoTalarLeñador As Byte = 2

Public Const EsfuerzoPescarPescador As Byte = 1
Public Const EsfuerzoPescarGeneral As Byte = 3

Public Const EsfuerzoExcavarMinero As Byte = 2
Public Const EsfuerzoExcavarGeneral As Byte = 5
Public Nombre1 As String
Public Nombre2 As String

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum PartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public Const Guardias As Integer = 6

Public Const MAXREP As Long = 6000000
Public Const MAXEXP As Long = 99999999

Public Const MAXATRIBUTOS As Byte = 38
Public Const MINATRIBUTOS As Byte = 6

Public Const LingoteHierro As Integer = 386
Public Const LingotePlata As Integer = 387
Public Const LingoteOro As Integer = 388
Public Const Leña As Integer = 58


Public Const MAXNPCS As Integer = 10000
Public Const MAXCHARS As Integer = 10000

Public Const HACHA_LEÑADOR As Integer = 127
Public Const PIQUETE_MINERO As Integer = 187

Public Const DAGA As Integer = 15
Public Const FOGATA_APAG As Integer = 136
Public Const FOGATA As Integer = 63
Public Const ORO_MINA As Integer = 194
Public Const PLATA_MINA As Integer = 193
Public Const HIERRO_MINA As Integer = 192
Public Const MARTILLO_HERRERO As Integer = 389
Public Const SERRUCHO_CARPINTERO As Integer = 198
Public Const RED_PESCA As Integer = 543
Public Const CAÑA_PESCA As Integer = 138

Public Enum eNPCType
    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Timbero = 7
    Guardiascaos = 8
    Viajero = 9
    ReyCastillo = 10 'By Nait
    Combinador = 11
    Viajerofer = 12
    Ciudadania = 13
End Enum

Public Const MapCastilloN = 31
Public Const MapCastilloS = 32
Public Const MapCastilloE = 33
Public Const MapCastilloO = 34 'By Nait

Public Const MIN_APUÑALAR As Byte = 10

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS As Byte = 22

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS As Byte = 5

''
' Cantidad de Clases
Public Const NUMCLASES As Byte = 17

''
' Cantidad de Razas
Public Const NUMRAZAS As Byte = 5


''
' Valor maximo de cada skill
Public Const MAXSKILLPOINTS As Byte = 100

''

''
'Direccion
'
' @param NORTH Norte
' @param EAST Este
' @param SOUTH Sur
' @param WEST Oeste
'
Public Enum eHeading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

''
' Cantidad maxima de mascotas
Public Const MAXMASCOTAS As Byte = 3

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO As Integer = 100
Public Const vlASESINO As Integer = 1000
Public Const vlCAZADOR As Integer = 5
Public Const vlNoble As Integer = 5
Public Const vlLadron As Integer = 25
Public Const vlProleta As Integer = 2

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto As Integer = 206
Public Const iCabezaMuerto As Integer = 512
Public Const iCuerpoMuertoCrimi = 205
Public Const iCabezaMuertoCrimi = 511
Public Const iCuerpoMuertoNeutro = 145
Public Const iCabezaMuertoNeutro = 501


Public Const iORO As Byte = 12
Public Const Pescado As Byte = 139

Public Enum PECES_POSIBLES
    PESCADO1 = 139
    PESCADO2 = 544
    PESCADO3 = 545
    PESCADO4 = 546
End Enum

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Enum eSkill
    Suerte = 1
    Magia = 2
    Robar = 3
    Tacticas = 4
    Armas = 5
    Meditar = 6
    Apuñalar = 7
    Ocultarse = 8
    Supervivencia = 9
    Talar = 10
    Comerciar = 11
    Defensa = 12
    Pesca = 13
    Mineria = 14
    Carpinteria = 15
    Herreria = 16
    Liderazgo = 17
    Domar = 18
    Proyectiles = 19
    Wresterling = 20
    Navegacion = 21
    Equitacion = 22
End Enum

Public Const FundirMetal = 88

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalHPCazador As Byte = 1 'HP adicionales cuando sube de nivel

Public Const AumentoSTDef As Byte = 15
Public Const AumentoSTLadron As Byte = AumentoSTDef + 3
Public Const AumentoSTMago As Byte = AumentoSTDef - 1
Public Const AumentoSTLeñador As Byte = AumentoSTDef + 23
Public Const AumentoSTPescador As Byte = AumentoSTDef + 20
Public Const AumentoSTMinero As Byte = AumentoSTDef + 25

'Tamaño del mapa
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'Tamaño del tileset
Public Const TileSizeX As Byte = 32
Public Const TileSizeY As Byte = 32

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow As Byte = 17
Public Const YWindow As Byte = 13

'Sonidos
Public Const SND_SWING As Byte = 2
Public Const SND_TALAR As Byte = 13
Public Const SND_PESCAR As Byte = 14
Public Const SND_MINERO As Byte = 15
Public Const SND_WARP As Byte = 3
Public Const SND_TRANSF As Byte = 58
Public Const SND_PUERTA As Byte = 5
Public Const SND_NIVEL As Byte = 6

Public Const SND_USERMUERTE As Byte = 11
Public Const SND_IMPACTO As Byte = 10
Public Const SND_IMPACTO2 As Byte = 12
Public Const SND_LEÑADOR As Byte = 13
Public Const SND_FOGATA As Byte = 14
Public Const SND_AVE As Byte = 21
Public Const SND_AVE2 As Byte = 22
Public Const SND_AVE3 As Byte = 34
Public Const SND_GRILLO As Byte = 28
Public Const SND_GRILLO2 As Byte = 29
Public Const SND_SACARARMA As Byte = 25
Public Const SND_ESCUDO As Byte = 37
Public Const MARTILLOHERRERO As Byte = 41
Public Const LABUROCARPINTERO As Byte = 42
Public Const SND_BEBER As Byte = 46

''
' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS As Integer = 10000

''
' Cantidad de "slots" en el inventario
Public Const MAX_INVENTORY_SLOTS As Byte = 20

' CATEGORIAS PRINCIPALES
Public Enum eOBJType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otESCUDO = 16
    otCASCO = 17
    otHerramientas = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otMonturas = 36
    otBolsaTesoro = 37
    otMapaTesoro = 38
    otCualquiera = 1000
End Enum

'Texto
Public Const FONTTYPE_TALK As String = "~255~255~255~0~0"
Public Const FONTTYPE_CELESTE_NEGRITA As String = "~0~128~255~1~0"
Public Const FONTTYPE_GANAR As String = "~240~240~50~1~0"
Public Const FONTTYPE_UDP As String = "~255~0~0~0~0"
Public Const FONTTYPE_FIGHT As String = "~255~0~0~1~0"
Public Const FONTTYPE_WARNING As String = "~32~51~223~1~1"
Public Const FONTTYPE_INFO As String = "~65~190~156~0~0"
Public Const FONTTYPE_INFOBOLD As String = "~65~190~156~1~0"
Public Const FONTTYPE_EJECUCION As String = "~130~130~130~1~0"
Public Const FONTTYPE_PARTY As String = "~255~180~255~0~0"
Public Const FONTTYPE_VENENO As String = "~0~255~0~0~0"
Public Const FONTTYPE_GUILD As String = "~255~255~255~1~0"
Public Const FONTTYPE_SERVER As String = "~0~185~0~0~0"
Public Const FONTTYPE_GUILDMSG As String = "~228~199~27~0~0"
Public Const FONTTYPE_CONSEJO As String = "~0~255~255~1~0"
Public Const FONTTYPE_CONSEJOCAOS As String = "~185~0~4~1~0"
Public Const FONTTYPE_CONSEJOVesA As String = "~0~255~255~1~0"
Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~185~0~4~1~0"
Public Const FONTTYPE_CENTINELA As String = "~0~255~0~1~0"
Public Const FONTTYPE_ADVERTENCIAS As String = "~128~0~0~1~1"

Public Const FONTTYPE_BLANCO As String = "~255~255~255~0~0"
Public Const FONTTYPE_BORDO As String = "~128~0~0~0~0"
Public Const FONTTYPE_VERDE As String = "~0~255~0~0~0"
Public Const FONTTYPE_ROJO As String = "~255~0~0~0~0"
Public Const FONTTYPE_AZUL As String = "~0~0~255~0~0"
Public Const FONTTYPE_VIOLETA As String = "~128~0~128~0~0"
Public Const FONTTYPE_AMARILLO As String = "~255~255~0~0~0"
Public Const FONTTYPE_CELESTE As String = "~128~255~255~0~0"
Public Const FONTTYPE_GRIS As String = "~130~130~130~0~0"

Public Const FONTTYPE_BLANCON As String = "~255~255~255~1~0"
Public Const FONTTYPE_BORDON As String = "~128~0~0~1~0"
Public Const FONTTYPE_VERDEN As String = "~0~255~0~1~0"
Public Const FONTTYPE_ROJON As String = "~255~0~0~1~0"
Public Const FONTTYPE_AZULN As String = "~0~0~255~1~0"
Public Const FONTTYPE_VIOLETAN As String = "~128~0~128~1~0"
Public Const FONTTYPE_AMARILLON As String = "~255~255~0~1~0"
Public Const FONTTYPE_CELESTEN As String = "~128~255~255~1~0"
Public Const FONTTYPE_GRISN As String = "~130~130~130~1~0"
Public Const FONTTYPE_ORO As String = "~255~255~0~1~0"

'Estadisticas
Public Const STAT_MAXELV As Byte = 50
Public Const STAT_MAXHP As Integer = 30000
Public Const STAT_MAXSTA As Integer = 30000
Public Const STAT_MAXMAN As Integer = 30000
Public Const STAT_MAXHIT_UNDER36 As Byte = 99
Public Const STAT_MAXHIT_OVER36 As Integer = 999
Public Const STAT_MAXDEF As Byte = 99



' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************

Public Type tHechizo
    Nombre As String
    Desc As String
    PalabrasMagicas As String
    ExclusivoClase As String
    ProhibidoClase As String
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
    Resis As Byte
    
    Tipo As TipoHechizo
    
    WAV As Integer
    FXgrh As Integer
    loops As Byte
    
    SubeHP As Byte
    MinHP As Integer
    MaxHP As Integer
    
    SubeMana As Byte
    MiMana As Integer
    MaMana As Integer
    
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
        
    Lenteja As Byte
    ActivaVIP As Byte
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer
    
    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    RemoverEstupidez As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Morph As Byte
    Mimetiza As Byte
    RemueveInvisibilidadParcial As Byte
    
    Invoca As Byte
    NumNpc As Integer
    Cant As Integer
    
    Materializa As Byte
    ItemIndex As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer

    'Barrin 29/9/03
    StaRequerido As Integer

    Target As TargetType
    
    NeedStaff As Integer
    StaffAffected As Boolean
End Type

Public Type LevelSkill
    LevelValue As Integer
End Type

Public Type UserOBJ
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
    ProbTirar As Byte
End Type

Public Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    HerramientaEqpObjIndex As Integer
    HerramientaEqpSlot As Integer
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    MonturaObjIndex As Integer
    MonturaSlot As Byte
    NroItems As Integer
End Type

Public Type tPartyData
    PIndex As Integer
    RemXP As Double 'La exp. en el server se cuenta con Doubles
    TargetUser As Integer 'Para las invitaciones
End Type

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type FXdata
    Nombre As String
    GrhIndex As Integer
    Delay As Integer
End Type

'Datos de user o npc
Public Type Char
    Aura As Integer
    CharIndex As Integer
    Head As Integer
    Body As Integer
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    FX As Integer
    loops As Integer
    
    Heading As eHeading
End Type

'Tipos de objetos
Public Type ObjData
    Aura As Integer
    name As String 'Nombre del obj
    
    OBJType As eOBJType 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    'Solo contenedores
    MaxItems As Integer
    Conte As Inventario
    Apuñala As Byte
    
    HechizoIndex As Integer
    DosManos As Byte
    
    ForoID As String
    
    MinHP As Integer ' Minimo puntos de vida
    MaxHP As Integer ' Maximo puntos de vida
    
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    
    proyectil As Integer
    Municion As Integer
    
    Crucial As Byte
    Newbie As Integer
    
    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina
    
    'Pociones
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHIT As Integer 'Minimo golpe
    MaxHIT As Integer 'Maximo golpe
    
    def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    
    Valor As Long     ' Precio
    Skill As Byte
    SkillM As Byte
    
    Cerrada As Integer
    Llave As Byte
    clave As Long 'si clave=llave la puerta se abre o cierra
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    RazaEnana As Byte
    Mujer As Byte
    Hombre As Byte
    
    Envenena As Byte
    Paraliza As Byte
    
    Agarrable As Byte
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    
    SkHerreria As Integer
    SkCarpinteria As Integer
    
    texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As String
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    Real As Integer
    Caos As Integer
    
    NoSeCae As Integer
    SoloVIP As Byte
    
    StaffPower As Integer
    StaffDamageBonus As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
'[/KEVIN]

'[KEVIN]
Public Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer
End Type
'[/KEVIN]


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type tReputacion 'Fama del usuario
    NobleRep As Double
    BurguesRep As Double
    PlebeRep As Double
    LadronesRep As Double
    BandidoRep As Double
    AsesinoRep As Double
    Promedio As Double
End Type

'Estadisticas de los usuarios
Public Type UserStats
    PuntosDonacion As Long
    Repu As Integer
    TransformadoVIP As Byte
    PuntosVIP As Integer
    RetosGanados As Long
    RetosPerdidos As Long
    DuelosGanados As Long
    DuelosPerdidos As Long
    PuntosTorneo As Long
    TrofOro As Byte
    MedOro As Byte
    TrofBronce As Byte
    TrofPlata As Byte
    MET As Integer
    
    MaxHP As Integer
    MinHP As Integer
    
    FIT As Integer
    MaxSta As Integer
    MinSta As Integer
    MaxMan As Integer
    MinMAN As Integer
    MaxHIT As Integer
    MinHIT As Integer
    
    MaxHam As Byte
    MinHam As Byte
    
    MaxAGU As Byte
    MinAGU As Byte
        
    def As Integer
    Exp As Double
    ELV As Long
    ELU As Long
    UserSkills(1 To NUMSKILLS) As Integer
    UserAtributos(1 To NUMATRIBUTOS) As Integer
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Integer
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UserQuests(1 To MAXUSERQUESTS) As tUserQuest
    UserQuestsDone As String
    UsuariosMatados As Integer
    CriminalesMatados As Integer
    NPCsMuertos As Integer
    
    SkillPts As Integer
    
End Type

Public Desafio As Desafio
 
Public Type Desafio
Primero As Integer
Segundo As Integer
End Type

Public Torne As Torne

Public Type Torne
Jugador1 As Integer
Jugador2 As Integer
Jugador3 As Integer
Jugador4 As Integer
Jugador5 As Integer
Jugador6 As Integer
Jugador7 As Integer
Jugador8 As Integer
Jugador9 As Integer
Jugador10 As Integer
Jugador11 As Integer
Jugador12 As Integer
Jugador13 As Integer
Jugador14 As Integer
Jugador15 As Integer
Jugador16 As Integer
Jugador17 As Integer
Jugador18 As Integer
Jugador19 As Integer
Jugador20 As Integer
Jugador21 As Integer
Jugador22 As Integer
Jugador23 As Integer
Jugador24 As Integer
Jugador25 As Integer
Jugador26 As Integer
Jugador27 As Integer
Jugador28 As Integer
Jugador29 As Integer
Jugador30 As Integer
Jugador31 As Integer
Jugador32 As Integer
End Type

Public Pareja As Pareja
 
Public Type Pareja
Jugador1 As Integer
Jugador2 As Integer
Jugador3 As Integer
Jugador4 As Integer
End Type

    Public dMap        As String
    Public dUser       As String
    Public dIndex      As String
    
'Status - Mithrandir
Public Type UserMithStatus
EsStatus As Byte
EligioStatus As Byte
End Type
'Status - Mithrandir

'Sistema de Consejo nuevo - Mithrandir
Public Type UserConsejos
PertAlCons As Byte
PertAlConsCaos As Byte
LiderConsejo As Byte
LiderConsejoCaos As Byte
End Type
'Sistema de Consejo nuevo - Mithrandir

'Flags
Public Type UserFlags
    EnTD As Byte
    TeamTD As Byte
    RondasDuelo As Byte
    DeathMatch As Byte
    YaestaJugando As Byte
    Stopped As Boolean
    Desenterrando As Byte
    Automaticop As Boolean
    ActivoGema As Byte
    GemaActivada As String
    TimeGema As Byte
    automatico As Boolean
    VIP As Byte
    ClienteValido As Integer
    TimeRevivir As Byte
    DueleandoTorneo As Boolean
    DueleandoTorneo2 As Boolean
    DueleandoTorneo3 As Boolean
    DueleandoTorneo4 As Boolean
    DueleandoFinal As Boolean
    DueleandoFinal2 As Boolean
    DueleandoFinal3 As Boolean
    DueleandoFinal4 As Boolean
    SeguroCVC As Boolean
    EnTorneo As Byte
    LeMandaronDuelo As Boolean
    UltimoEnMandarDuelo As String
    EnDuelo As Boolean
    DueliandoContra As String
    PJerarquia As Byte
    SJerarquia As Byte
    TJerarquia As Byte
    CJerarquia As Byte
    CJerarquiaC As Byte
    SuPareja As Integer
    EsperaPareja As Boolean
    EnPareja As Boolean
    Desafio As Integer
    EnDesafio As Integer
    rondas As Integer
    EsperandoDuelo As Boolean
    EstaDueleando As Boolean
    Oponente As Integer
    EsperandoDueloxset As Boolean
    EstaDueleandoxset As Boolean
    Oponentexset As Integer
    DeseoRecibirMSJ As Byte
    EstaEmpo As Byte    'Empollando (by yb)
    EleDeTierra As Byte
    EleDeFuego As Byte
    EleDeAgua As Byte
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    Meditando As Boolean
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    
    Vuela As Byte
    Navegando As Byte
    Montando As Byte
    Transformado As Byte
    Seguro As Boolean
    SeguroResu As Boolean
    SeguroClan As Boolean
    
    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As eNPCType ' Tipo del npc señalado
    NpcInv As Integer
    
    Ban As Byte
    AdministrativeBan As Byte
    
    TargetUser As Integer ' Usuario señalado
    
    TargetObj As Integer ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
    
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
    
    StatsChanged As Byte
    Privilegios As PlayerType
    EsRolesMaster As Boolean
    
    MalWaWe As Integer
    
    LastCrimMatado As String
    LastCiudMatado As String
    LastNeutrMatado As String
    
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte

    '[CDT 17-02-04]
    UltimoMensaje As Byte
    '[/CDT]
       NoActualizado As Boolean
    PertAlCons As Byte
    PertAlConsCaos As Byte
    
    Silenciado As Byte
    
    Mimetizado As Byte
    
    CentinelaOK As Boolean 'Centinela
End Type

Public Type UserCounters
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Veneno As Integer
    Paralisis As Integer
    Ceguera As Integer
    Estupidez As Integer
    Invisibilidad As Integer
    Mimetismo As Integer
    PiqueteC As Long
    Pena As Long
    SendMapCounter As WorldPos
    Pasos As Integer
    TiraItem As Byte
    TimeComandos As Byte
    LentejaTiempo As Byte
    MuereEnTD As Byte
    EntreTiempo As Byte
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]
    
    'Barrin 3/10/03
    tInicioMeditar As Long
    bPuedeMeditar As Boolean
    'Barrin
    
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    
    Trabajando As Long  ' Para el centinela
    Ocultando As Long   ' Unico trabajo no revisado por el centinela
End Type

Public Type tFacciones
    ArmadaReal As Byte
    FuerzasCaos As Byte
    CriminalesMatados As Double
    CiudadanosMatados As Double
    NeutralesMatados As Double
    RecompensasReal As Long
    RecompensasCaos As Long
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
    Reenlistadas As Byte
End Type

'Tipo de los Usuarios
Public Type User

    Lac As TLac '[loopzer] 'el Anti-Cheats Lac(Loopzer Anti-Cheats)
    
    name As String
    ID As Long
    
    showName As Boolean 'Permite que los GMs oculten su nick con el comando /SHOWNAME
    
    modName As String
    PassWord As String
    
    Char As Char 'Define la apariencia
    CharMimetizado As Char
    OrigChar As Char
    
    Desc As String ' Descripcion
    DescRM As String
    
    Clase As String
    Raza As String
    Genero As String
    email As String
    Pin As String
    Hogar As String
        
    Invent As Inventario
    
    pos As WorldPos
    
    ConnIDValida As Boolean
    ConnID As Long 'ID
    RDBuffer As String 'Buffer roto
    
    CommandsBuffer As New CColaArray
    ColaSalida As New Collection
    SockPuedoEnviar As Boolean
    
    '[KEVIN]
    BancoInvent As BancoInventario
    '[/KEVIN]
    
    Counters As UserCounters
    
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMacotas As Integer
    
    Stats As UserStats
    flags As UserFlags
    ConsejoInfo As UserConsejos
    StatusMith As UserMithStatus
    NumeroPaquetesPorMiliSec As Long
    BytesTransmitidosUser As Long
    BytesTransmitidosSvr As Long
    
    Reputacion As tReputacion
    
    Faccion As tFacciones
    
    PrevCheckSum As Long
    PacketNumber As Long
    RandKey As Long
    
    ip As String
    hd As String
    EnCvc As Boolean
     '[Alejo]
    ComUsu As tCOmercioUsuario
    '[/Alejo]
    
    EmpoCont As Byte
    
    GuildIndex As Integer   'puntero al array global de guilds
    FundandoGuildAlineacion As ALINEACION_GUILD     'esto esta aca hasta que se parchee el cliente y se pongan cadenas de datos distintas para cada alineacion
    EscucheClan As Integer
    
    PartyIndex As Integer   'index a la party q es miembro
    PartySolicitud As Integer   'index a la party q solicito
    
    KeyCrypt As Integer
    
    AreasInfo As AreaInfo
End Type


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type NPCStats
    Alineacion As Integer
    MaxHP As Long
    MinHP As Long
    MaxHIT As Integer
    MinHIT As Integer
    def As Integer
    UsuariosMatados As Integer
End Type

Public Type NpcCounters
    Paralisis As Integer
    TiempoExistencia As Long
End Type

Public Type NPCFlags
    AfectaParalisis As Byte
    GolpeExacto As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    Faccion As Byte
    LanzaSpells As Byte
    
    '[KEVIN]
    'DeQuest As Byte
    
    'ExpDada As Long
    ExpCount As Long '[ALEJO]
    '[/KEVIN]
    
    OldMovement As TipoAI
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    UseAINow As Boolean
    Sound As Integer
    Attacking As Integer
    AttackedBy As String
    Category1 As String
    Category2 As String
    Category3 As String
    Category4 As String
    Category5 As String
    BackUp As Byte
    RespawnOrigPos As Byte
    
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    AtacaAPJ As Integer
    AtacaANPC As Integer
    AIAlineacion As e_Alineacion
    AIPersonalidad As e_Personalidad
End Type

Public Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

' New type for holding the pathfinding info
Public Type NpcPathFindingInfo
    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    TargetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location
    
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
    
End Type
' New type for holding the pathfinding info


Public Type npc
    name As String
    Char As Char 'Define como se vera
    Desc As String
    DescExtra As String

    NPCtype As eNPCType
    Numero As Integer

    level As Integer

    InvReSpawn As Byte

    Comercia As Integer
    Target As Long
    TargetNPC As Long
    TipoItems As Integer

    Veneno As Byte

    pos As WorldPos 'Posicion
    Orig As WorldPos
    SkillDomar As Integer

    Movement As TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long
    Aura As Integer

    GiveEXP As Long
    GivePT As Long

    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    
    Invent As Inventario
    CanAttack As Byte
    
    NroExpresiones As Byte
    Expresiones() As String ' le da vida ;)
    
    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
    
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer
    
    ' New!! Needed for pathfindig
    PFINFO As NpcPathFindingInfo
    AreasInfo As AreaInfo
    QuestNumber As Integer
    TalkAfterQuest As String
    TalkDuringQuest As String
End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Public Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Integer
    userindex As Integer
    NpcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    trigger As eTrigger
End Type

Type MoveeTo
Activo As Byte
Map As Integer
X As Integer
Y As Integer
End Type

'Info del mapa
Type MapInfo
    SeCaenItems As Byte
    Transport As MoveeTo
    criatinv As Integer
    NumUsers As Integer
    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer
    Pk As Boolean
    MagiaSinEfecto As Byte
    NoEncriptarMP As Byte
    
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
End Type

'********** V A R I A B L E S     P U B L I C A S ***********

Public SERVERONLINE As Boolean
Public ULTIMAVERSION As String
Public BackUp As Boolean ' TODO: Se usa esta variable ?

Public ListaRazas(1 To NUMRAZAS) As String
Public SkillsNames(1 To NUMSKILLS) As String
Public ListaClases(1 To NUMCLASES) As String

Public Const ENDL As String * 2 = vbCrLf
Public Const ENDC As String * 1 = vbNullChar

Public recordusuarios As Long

'
'Directorios
'

''
'Ruta base del server, en donde esta el "server.ini"
Public IniPath As String


''
'Ruta base para guardar los chars
Public CharPath As String

''
'Ruta base para los archivos de mapas
Public MapPath As String

''
'Ruta base para los DATs
Public DatPath As String

''
'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public ResPos As WorldPos ' TODO: Se usa esta variable ?

''
'Posicion de comienzo
Public StartPos As WorldPos ' TODO: Se usa esta variable ?

''
'Numero de usuarios actual
Public NumUsers As Integer
Public LastUser As Integer
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumFX As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public NumeroHechizos As Integer
Public AllowMultiLogins As Byte
Public IdleLimit As Integer
Public MaxUsers As Integer
Public HideMe As Byte
Public LastBackup As String
Public Minutos As String
Public haciendoBK As Boolean

Public Torneo_Map As Integer
Public Torneo_X As Byte
Public Torneo_Y As Byte
Public Torneo_Nivel_Minimo As Byte
Public Torneo_Nivel_Maximo As Byte
Public Torneo_Cantidad As Byte
Public Torneo_SumAuto As Byte
Public Hay_Torneo As Boolean
Public Torneo As New cCola
Public Torneo_Clases_Validas(1 To 8) As String
Public Torneo_Clases_Validas2(1 To 8) As Byte
Public Torneo_Alineacion_Validas(1 To 4) As String
Public Torneo_Alineacion_Validas2(1 To 4) As Byte

Public PuedeCrearPersonajes As Integer
Public CamaraLenta As Integer
Public ServerSoloGMs As Integer
'Esta activada la verificacion MD5 ?
Public MD5ClientesActivado As Byte

Public EnPausa As Boolean
Public EnTesting As Boolean
Public EncriptarProtocolosCriticos As Boolean

'*****************ARRAYS PUBLICOS*************************
Public UserList() As User 'USUARIOS
Public Npclist() As npc 'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList() As Integer
Public ObjData() As ObjData
Public FX() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To 50) As LevelSkill
Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ObjCarpintero() As Integer
Public MD5s() As String
Public BanIps As New Collection
Public Parties() As clsParty
'*********************************************************

Public Helkat As WorldPos
Public Runek As WorldPos
Public Banderbill As WorldPos
Public Lindos As WorldPos

Public Prision As WorldPos
Public Libertad As WorldPos

Public Ayuda As New cCola
Public ConsultaPopular As New ConsultasPopulares
Public SonidosMapas As New SoundMapInfo

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Enum e_ObjetosCriticos
    Manzana = 1
    Manzana2 = 2
    ManzanaNewbie = 467
End Enum
