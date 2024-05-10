Attribute VB_Name = "SistemaCombate"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Diseño y corrección del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18

Function ModificadorEvasion(ByVal Clase As String) As Single

Select Case UCase$(Clase)
    Case "GUERRERO"
        ModificadorEvasion = 0.9
    Case "CAZADOR"
        ModificadorEvasion = 0.9
    Case "PALADIN"
        ModificadorEvasion = 0.85
    Case "BANDIDO"
        ModificadorEvasion = 0.9
    Case "ASESINO"
        ModificadorEvasion = 0.9
    Case "PIRATA"
        ModificadorEvasion = 0.9
    Case "LADRON"
        ModificadorEvasion = 1.1
    Case "BARDO"
        ModificadorEvasion = 0.88
    Case "CLERIGO"
        ModificadorEvasion = 0.81
    Case "MAGO"
        ModificadorEvasion = 0.63
    Case "DRUIDA"
        ModificadorEvasion = 0.7
    Case Else
        ModificadorEvasion = 0.8
End Select
End Function

Function ModificadorPoderAtaqueArmas(ByVal Clase As String) As Single
Select Case UCase$(Clase)
    Case "GUERRERO"
        ModificadorPoderAtaqueArmas = 0.9
    Case "CAZADOR"
        ModificadorPoderAtaqueArmas = 0.75
    Case "PALADIN"
        ModificadorPoderAtaqueArmas = 0.85
    Case "ASESINO"
        ModificadorPoderAtaqueArmas = 0.82
    Case "PIRATA"
        ModificadorPoderAtaqueArmas = 0.8
    Case "LADRON"
        ModificadorPoderAtaqueArmas = 0.75
    Case "BANDIDO"
        ModificadorPoderAtaqueArmas = 0.75
    Case "CLERIGO"
        ModificadorPoderAtaqueArmas = 0.9
    Case "BARDO"
        ModificadorPoderAtaqueArmas = 0.75
    Case "DRUIDA"
        ModificadorPoderAtaqueArmas = 0.7
    Case "PESCADOR"
        ModificadorPoderAtaqueArmas = 0.6
    Case "LEÑADOR"
        ModificadorPoderAtaqueArmas = 0.6
    Case "MINERO"
        ModificadorPoderAtaqueArmas = 0.6
    Case "HERRERO"
        ModificadorPoderAtaqueArmas = 0.6
    Case "CARPINTERO"
        ModificadorPoderAtaqueArmas = 0.6
    Case Else
        ModificadorPoderAtaqueArmas = 0.5
End Select
End Function

Function ModificadorPoderAtaqueProyectiles(ByVal Clase As String) As Single
Select Case UCase$(Clase)
    Case "GUERRERO"
        ModificadorPoderAtaqueProyectiles = 0.85
    Case "CAZADOR"
        ModificadorPoderAtaqueProyectiles = 1
    Case "PALADIN"
        ModificadorPoderAtaqueProyectiles = 0.75
    Case "ASESINO"
        ModificadorPoderAtaqueProyectiles = 0.75
    Case "PIRATA"
        ModificadorPoderAtaqueProyectiles = 0.75
    Case "LADRON"
        ModificadorPoderAtaqueProyectiles = 0.8
    Case "BANDIDO"
        ModificadorPoderAtaqueProyectiles = 0.8
    Case "CLERIGO"
        ModificadorPoderAtaqueProyectiles = 0.7
    Case "BARDO"
        ModificadorPoderAtaqueProyectiles = 0.7
    Case "DRUIDA"
        ModificadorPoderAtaqueProyectiles = 0.75
    Case "PESCADOR"
        ModificadorPoderAtaqueProyectiles = 0.65
    Case "LEÑADOR"
        ModificadorPoderAtaqueProyectiles = 0.7
    Case "MINERO"
        ModificadorPoderAtaqueProyectiles = 0.65
    Case "HERRERO"
        ModificadorPoderAtaqueProyectiles = 0.65
    Case "CARPINTERO"
        ModificadorPoderAtaqueProyectiles = 0.7
    Case Else
        ModificadorPoderAtaqueProyectiles = 0.5
End Select
End Function

Function ModicadorDañoClaseArmas(ByVal Clase As String) As Single
Select Case UCase$(Clase)
    Case "GUERRERO"
        ModicadorDañoClaseArmas = 0.95
    Case "CAZADOR"
        ModicadorDañoClaseArmas = 0.8
    Case "PALADIN"
        ModicadorDañoClaseArmas = 0.88
    Case "ASESINO"
        ModicadorDañoClaseArmas = 0.9
    Case "LADRON"
        ModicadorDañoClaseArmas = 0.8
    Case "PIRATA"
        ModicadorDañoClaseArmas = 0.8
    Case "BANDIDO"
        ModicadorDañoClaseArmas = 0.8
    Case "CLERIGO"
        ModicadorDañoClaseArmas = 0.85
    Case "BARDO"
        ModicadorDañoClaseArmas = 0.75
    Case "DRUIDA"
        ModicadorDañoClaseArmas = 0.75
    Case "PESCADOR"
        ModicadorDañoClaseArmas = 0.6
    Case "LEÑADOR"
        ModicadorDañoClaseArmas = 0.7
    Case "MINERO"
        ModicadorDañoClaseArmas = 0.75
    Case "HERRERO"
        ModicadorDañoClaseArmas = 0.75
    Case "CARPINTERO"
        ModicadorDañoClaseArmas = 0.7
    Case Else
        ModicadorDañoClaseArmas = 0.5
End Select
End Function

Function ModicadorDañoClaseProyectiles(ByVal Clase As String) As Single
Select Case UCase$(Clase)
    Case "GUERRERO"
        ModicadorDañoClaseProyectiles = 1
    Case "CAZADOR"
        ModicadorDañoClaseProyectiles = 1.05
    Case "PALADIN"
        ModicadorDañoClaseProyectiles = 0.8
    Case "ASESINO"
        ModicadorDañoClaseProyectiles = 0.8
    Case "LADRON"
        ModicadorDañoClaseProyectiles = 0.75
    Case "PIRATA"
        ModicadorDañoClaseProyectiles = 0.75
    Case "BANDIDO"
        ModicadorDañoClaseProyectiles = 0.75
    Case "CLERIGO"
        ModicadorDañoClaseProyectiles = 0.7
    Case "BARDO"
        ModicadorDañoClaseProyectiles = 0.7
    Case "DRUIDA"
        ModicadorDañoClaseProyectiles = 0.75
    Case "PESCADOR"
        ModicadorDañoClaseProyectiles = 0.6
    Case "LEÑADOR"
        ModicadorDañoClaseProyectiles = 0.7
    Case "MINERO"
        ModicadorDañoClaseProyectiles = 0.6
    Case "HERRERO"
        ModicadorDañoClaseProyectiles = 0.6
    Case "CARPINTERO"
        ModicadorDañoClaseProyectiles = 0.7
    Case Else
        ModicadorDañoClaseProyectiles = 0.5
End Select
End Function

Function ModEvasionDeEscudoClase(ByVal Clase As String) As Single

Select Case UCase$(Clase)
Case "GUERRERO"
        ModEvasionDeEscudoClase = 0.9
    Case "CAZADOR"
        ModEvasionDeEscudoClase = 0.8
    Case "PALADIN"
        ModEvasionDeEscudoClase = 0.96
    Case "ASESINO"
        ModEvasionDeEscudoClase = 0.79
    Case "LADRON"
        ModEvasionDeEscudoClase = 0.7
    Case "BANDIDO"
        ModEvasionDeEscudoClase = 0.8
    Case "PIRATA"
        ModEvasionDeEscudoClase = 0.75
    Case "CLERIGO"
        ModEvasionDeEscudoClase = 0.79
    Case "BARDO"
        ModEvasionDeEscudoClase = 0.8
    Case "DRUIDA"
        ModEvasionDeEscudoClase = 0.75
    Case "PESCADOR"
        ModEvasionDeEscudoClase = 0.7
    Case "LEÑADOR"
        ModEvasionDeEscudoClase = 0.7
    Case "MINERO"
        ModEvasionDeEscudoClase = 0.7
    Case "HERRERO"
        ModEvasionDeEscudoClase = 0.7
    Case "CARPINTERO"
        ModEvasionDeEscudoClase = 0.7
    Case Else
        ModEvasionDeEscudoClase = 0.6
End Select

End Function
Function Minimo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Minimo = b
    Else: Minimo = a
End If
End Function

Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MinimoInt = b
    Else: MinimoInt = a
End If
End Function

Function Maximo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Maximo = a
    Else: Maximo = b
End If
End Function

Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MaximoInt = a
    Else: MaximoInt = b
End If
End Function


Function PoderEvasionEscudo(ByVal userindex As Integer) As Long

PoderEvasionEscudo = (UserList(userindex).Stats.UserSkills(eSkill.Defensa) * _
ModEvasionDeEscudoClase(UserList(userindex).Clase)) / 2

End Function

Function PoderEvasion(ByVal userindex As Integer) As Long
    Dim lTemp As Long
     With UserList(userindex)
       lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * _
          ModificadorEvasion(.Clase)
       
        PoderEvasion = (lTemp + (2.5 * Maximo(.Stats.ELV - 12, 0)))
    End With
End Function



'Function PoderEvasion(ByVal UserIndex As Integer) As Long
'Dim PoderEvasionTemp As Long

'If UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) < 31 Then
'    PoderEvasionTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) * _
'    ModificadorEvasion(UserList(UserIndex).Clase))
'ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) < 61 Then
'        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) + _
'        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
'        ModificadorEvasion(UserList(UserIndex).Clase))
'ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) < 91 Then
'        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) + _
'        (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
'        ModificadorEvasion(UserList(UserIndex).Clase))
'Else
'        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) + _
'        (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
'        ModificadorEvasion(UserList(UserIndex).Clase))
'End If
'PoderEvasion = (PoderEvasionTemp + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)))
'
'End Function
'
'
'



Function PoderAtaqueArma(ByVal userindex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(userindex).Stats.UserSkills(eSkill.Armas) < 31 Then
    PoderAtaqueTemp = (UserList(userindex).Stats.UserSkills(eSkill.Armas) * _
    ModificadorPoderAtaqueArmas(UserList(userindex).Clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Armas) < 61 Then
    PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Armas) + _
    UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad)) * _
    ModificadorPoderAtaqueArmas(UserList(userindex).Clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Armas) < 91 Then
    PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Armas) + _
    (2 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
    ModificadorPoderAtaqueArmas(UserList(userindex).Clase))
Else
   PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Armas) + _
   (3 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
   ModificadorPoderAtaqueArmas(UserList(userindex).Clase))
End If

PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * Maximo(UserList(userindex).Stats.ELV - 12, 0)))
End Function

Function PoderAtaqueProyectil(ByVal userindex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) < 31 Then
    PoderAtaqueTemp = (UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) * _
    ModificadorPoderAtaqueProyectiles(UserList(userindex).Clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) < 61 Then
        PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) + _
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueProyectiles(UserList(userindex).Clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) < 91 Then
        PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) + _
        (2 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueProyectiles(UserList(userindex).Clase))
Else
       PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Proyectiles) + _
      (3 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
      ModificadorPoderAtaqueProyectiles(UserList(userindex).Clase))
End If

PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * Maximo(UserList(userindex).Stats.ELV - 12, 0)))

End Function

Function PoderAtaqueWresterling(ByVal userindex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(userindex).Stats.UserSkills(eSkill.Wresterling) < 31 Then
    PoderAtaqueTemp = (UserList(userindex).Stats.UserSkills(eSkill.Wresterling) * _
    ModificadorPoderAtaqueArmas(UserList(userindex).Clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) < 61 Then
        PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Wresterling) + _
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueArmas(UserList(userindex).Clase))
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) < 91 Then
        PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Wresterling) + _
        (2 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueArmas(UserList(userindex).Clase))
Else
       PoderAtaqueTemp = ((UserList(userindex).Stats.UserSkills(eSkill.Wresterling) + _
       (3 * UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))) * _
       ModificadorPoderAtaqueArmas(UserList(userindex).Clase))
End If

PoderAtaqueWresterling = (PoderAtaqueTemp + (2.5 * Maximo(UserList(userindex).Stats.ELV - 12, 0)))

End Function


Public Function UserImpactoNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer) As Boolean
Dim PoderAtaque As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim ProbExito As Long

Arma = UserList(userindex).Invent.WeaponEqpObjIndex
If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

If Arma > 0 Then 'Usando un arma
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(userindex)
    Else
        PoderAtaque = PoderAtaqueArma(userindex)
    End If
Else 'Peleando con puños
    PoderAtaque = PoderAtaqueWresterling(userindex)
End If


ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))

UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

If UserImpactoNpc Then
    If Arma <> 0 Then
       If proyectil Then
            Call SubirSkill(userindex, Proyectiles)
       Else
            Call SubirSkill(userindex, Armas)
       End If
    Else
        Call SubirSkill(userindex, Wresterling)
    End If
End If


End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
'*************************************************
Dim Rechazo As Boolean
Dim ProbRechazo As Long
Dim ProbExito As Long
Dim UserEvasion As Long
Dim NpcPoderAtaque As Long
Dim PoderEvasioEscudo As Long
Dim SkillTacticas As Long
Dim SkillDefensa As Long

UserEvasion = PoderEvasion(userindex)
NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
PoderEvasioEscudo = PoderEvasionEscudo(userindex)

SkillTacticas = UserList(userindex).Stats.UserSkills(eSkill.Tacticas)
SkillDefensa = UserList(userindex).Stats.UserSkills(eSkill.Defensa)

'Esta usando un escudo ???
If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))

NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
    If Not NpcImpacto Then
        If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
            ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_ESCUDO)
                Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "EM" & UserList(userindex).Char.CharIndex)
                Call SendData(SendTarget.toindex, userindex, 0, "7")
                Call SubirSkill(userindex, Defensa)
            End If
        End If
    End If
End If
End Function


Public Function CalcularDaño(ByVal userindex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
Dim DañoArma As Long, DañoUsuario As Long, Arma As ObjData, ModifClase As Single
Dim proyectil As ObjData
Dim DañoMaxArma As Long
''sacar esto si no queremos q la matadracos mate el dragon si o si
Dim matodragon As Boolean
matodragon = False
 
 
If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
    Arma = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex)
   
   
    ' Ataca a un npc?
    If NpcIndex > 0 Then
       
        'Usa la mata dragones?
        If UserList(userindex).Invent.WeaponEqpObjIndex = 1053 And Npclist(NpcIndex).NPCtype = DRAGON Then ' Usa la matadragones?
            ModifClase = ModicadorDañoClaseArmas(UserList(userindex).Clase)
                DañoArma = RandomNumber(400, 500)
                DañoMaxArma = 500
        Else ' daño comun
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(userindex).Clase)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT) * 1.4
                DañoMaxArma = Arma.MaxHIT * 1.4
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(userindex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT) * 1.4
                    DañoMaxArma = Arma.MaxHIT * 1.4
                End If
           Else
                ModifClase = ModicadorDañoClaseArmas(UserList(userindex).Clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT) * 1.4
                DañoMaxArma = Arma.MaxHIT * 1.4
               
           End If
        End If
   
    Else ' Ataca usuario
        If UserList(userindex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
            ModifClase = ModicadorDañoClaseArmas(UserList(userindex).Clase)
                DañoArma = 1 ' Si usa la espada matadragones daño es 1
            DañoMaxArma = 1
        Else
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(userindex).Clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(userindex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDañoClaseArmas(UserList(userindex).Clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    End If
Else
    CalcularDaño = CInt(UserList(userindex).Stats.MaxHIT / 5)
    Exit Function
End If
 
DañoUsuario = RandomNumber(UserList(userindex).Stats.MinHIT, UserList(userindex).Stats.MaxHIT)
 
''sacar esto si no queremos q la matadracos mate el dragon si o si
If matodragon Then
    CalcularDaño = Npclist(NpcIndex).Stats.MinHP + Npclist(NpcIndex).Stats.def
Else
    CalcularDaño = (((3 * DañoArma) + ((DañoMaxArma / 5) * Maximo(0, (UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) - 15))) + DañoUsuario) * ModifClase)
End If
End Function

Public Sub UserDañoNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer)

If Npclist(NpcIndex).pos.Map = 106 And Npclist(NpcIndex).Numero = 937 And GuardiasRey <= 3 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar al rey mientras sus guerreros sigan con vida!." & FONTTYPE_ORO)
Exit Sub
End If

Dim daño As Long
 
 
 
daño = CalcularDaño(userindex, NpcIndex)
 
'esta navegando? si es asi le sumamos el daño del barco
If UserList(userindex).flags.Navegando = 1 Then _
        daño = daño + RandomNumber(ObjData(UserList(userindex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(userindex).Invent.BarcoObjIndex).MaxHIT)
        
If UserList(userindex).flags.Montando = 1 Then daño = daño + RandomNumber(ObjData(UserList(userindex).Invent.MonturaObjIndex).MinHIT, ObjData(UserList(userindex).Invent.MonturaObjIndex).MaxHIT)
 
daño = daño - Npclist(NpcIndex).Stats.def
 
If daño < 0 Then daño = 0
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbYellow & "°-" & daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))

'[KEVIN]
Call SendData(SendTarget.toindex, userindex, 0, "U2" & daño)
Call CheckPets(NpcIndex, userindex, False)
Call CalcularDarExp(userindex, NpcIndex, daño)
If UserList(userindex).flags.Privilegios = PlayerType.Admin Then
Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - Npclist(NpcIndex).Stats.MaxHP
Else
Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
End If
'[/KEVIN]
 
If Npclist(NpcIndex).Stats.MinHP > 0 Then
    'Trata de apuñalar por la espalda al enemigo
    If PuedeApuñalar(userindex) Then
       Call DoApuñalar(userindex, NpcIndex, 0, daño)
       Call SubirSkill(userindex, Apuñalar)
    End If
End If
 
 
If Npclist(NpcIndex).Stats.MinHP <= 0 Then

If Npclist(NpcIndex).NPCtype = DRAGON Then
            'Si tiene equipada la matadracos se la sacamos
            If UserList(userindex).Invent.WeaponEqpObjIndex = EspadaMataDragones Then
                Call QuitarObjetos(EspadaMataDragones, 1, userindex)
            End If
            If Npclist(NpcIndex).Stats.MaxHP > 100000 Then Call LogDesarrollo(UserList(userindex).name & " mató un dragón")
        End If
        
     ' Si era un Dragon perdemos la espada matadragones
        If Npclist(NpcIndex).NPCtype = DRAGON Then
            'Si tiene equipada la matadracos se la sacamos
            If UserList(userindex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                Call QuitarObjetos(EspadaMataDragonesIndex, 1, userindex)
            End If
            If Npclist(NpcIndex).Stats.MaxHP > 100000 Then Call LogDesarrollo(UserList(userindex).name & " mató un dragón")
        End If
       
       
        ' Para que las mascotas no sigan intentando luchar y
        ' comiencen a seguir al amo
       
        Dim j As Integer
        For j = 1 To MAXMASCOTAS
            If UserList(userindex).MascotasIndex(j) > 0 Then
                If Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = NpcIndex Then Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = 0
                Npclist(UserList(userindex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
            End If
        Next j
       
        Call MuereNpc(NpcIndex, userindex)
End If
 
End Sub


Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal userindex As Integer)
 
Dim daño As Integer, Lugar As Integer, absorbido As Integer, npcfile As String
Dim antdaño As Integer, defbarco As Integer, defMontura As Integer
Dim Obj As ObjData
 
 
 
daño = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
antdaño = daño
 
If UserList(userindex).flags.Navegando = 1 Then
    Obj = ObjData(UserList(userindex).Invent.BarcoObjIndex)
    defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

If UserList(userindex).flags.Montando = 1 Then
    Obj = ObjData(UserList(userindex).Invent.MonturaObjIndex)
    defMontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If
 
 
Lugar = RandomNumber(1, 6)
 

Select Case Lugar
  Case PartesCuerpo.bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(userindex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defMontura
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 1 Then daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
           Obj = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defMontura
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 1 Then daño = 1
        End If
End Select
 
If daño > 149 Then
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & &HFFFF& & "°" & "- " & daño & "" & "°" & str(UserList(userindex).Char.CharIndex))
Else
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & &HFFFF& & "°" & "- " & daño & "" & "°" & str(UserList(userindex).Char.CharIndex))
End If

Call SendData(SendTarget.toindex, userindex, 0, "N2" & Lugar & "," & daño)

If UserList(userindex).flags.Privilegios <= PlayerType.VIP Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - daño
 
'Muere el usuario
If UserList(userindex).Stats.MinHP <= 0 Then
 
    Call SendData(SendTarget.toindex, userindex, 0, "6") ' Le informamos que ha muerto ;)
   
    'Si lo mato un guardia
    If Criminal(userindex) And Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        Call RestarCriminalidad(userindex)
        If Not Criminal(userindex) And UserList(userindex).Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(userindex)
    End If
   
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
    Else
        'Al matarlo no lo sigue mas
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                    Npclist(NpcIndex).flags.AttackedBy = ""
        End If
    End If
   
    Call UserDie(userindex)
 
End If
 
End Sub

Public Sub RestarCriminalidad(ByVal userindex As Integer)
    'If UserList(UserIndex).Reputacion.AsesinoRep > 0 Then
    '     UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep - vlASESINO
    '     If UserList(UserIndex).Reputacion.AsesinoRep < 0 Then UserList(UserIndex).Reputacion.AsesinoRep = 0
    'Else
    If UserList(userindex).Reputacion.BandidoRep > 0 Then
         UserList(userindex).Reputacion.BandidoRep = UserList(userindex).Reputacion.BandidoRep - vlASALTO
         If UserList(userindex).Reputacion.BandidoRep < 0 Then UserList(userindex).Reputacion.BandidoRep = 0
    ElseIf UserList(userindex).Reputacion.LadronesRep > 0 Then
         UserList(userindex).Reputacion.LadronesRep = UserList(userindex).Reputacion.LadronesRep - (vlCAZADOR * 10)
         If UserList(userindex).Reputacion.LadronesRep < 0 Then UserList(userindex).Reputacion.LadronesRep = 0
    End If
End Sub


Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal userindex As Integer, Optional ByVal CheckElementales As Boolean = True)

Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(j) > 0 Then
       If UserList(userindex).MascotasIndex(j) <> NpcIndex Then
        If CheckElementales Or (Npclist(UserList(userindex).MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(UserList(userindex).MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
            If Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = 0 Then Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = NpcIndex
            'Npclist(UserList(UserIndex).MascotasIndex(j)).Flags.OldMovement = Npclist(UserList(UserIndex).MascotasIndex(j)).Movement
            Npclist(UserList(userindex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
        End If
       End If
    End If
Next j

End Sub
Public Sub AllFollowAmo(ByVal userindex As Integer)
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(j) > 0 Then
        Call FollowAmo(UserList(userindex).MascotasIndex(j))
    End If
Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean
 
If UserList(userindex).flags.AdminInvisible = 1 Then Exit Function
 
' El npc puede atacar ???
If Npclist(NpcIndex).CanAttack = 1 Then
    NpcAtacaUser = True
    Call CheckPets(NpcIndex, userindex, False)
 
    If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = userindex
 
    If UserList(userindex).flags.AtacadoPorNpc = 0 And _
       UserList(userindex).flags.AtacadoPorUser = 0 Then UserList(userindex).flags.AtacadoPorNpc = NpcIndex
Else
    NpcAtacaUser = False
    Exit Function
End If
 
Npclist(NpcIndex).CanAttack = 0
 
If Npclist(NpcIndex).flags.Snd1 > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Npclist(NpcIndex).flags.Snd1)
 
If NpcImpacto(NpcIndex, userindex) Then
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_IMPACTO)
   
    If UserList(userindex).flags.Meditando = False Then
        If UserList(userindex).flags.Navegando = 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXSANGRE & "," & 0)
        If UserList(userindex).flags.Navegando = 0 And UserList(userindex).flags.Montando = 0 Then Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXSANGRE & "," & 0)
    End If
   
    Call NpcDaño(NpcIndex, userindex)
    Call SendData(SendTarget.toindex, userindex, 0, "ASH" & UserList(userindex).Stats.MinHP)
    '¿Puede envenenar?
    If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(userindex)
Else
    Call SendData(SendTarget.toindex, userindex, 0, "N1")
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & &HFFFF& & "°¡Fallo!" & "°" & str(UserList(userindex).Char.CharIndex))
End If
 
 
 
'-----Tal vez suba los skills------
Call SubirSkill(userindex, Tacticas)
 
'Call SendUserStatsBox(UserIndex)
'Controla el nivel del usuario
Call CheckUserLevel(userindex)
 
End Function
Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
Dim PoderAtt As Long, PoderEva As Long, dif As Long
Dim ProbExito As Long

PoderAtt = Npclist(Atacante).PoderAtaque
PoderEva = Npclist(Victima).PoderEvasion
ProbExito = Maximo(10, Minimo(90, 50 + _
            ((PoderAtt - PoderEva) * 0.4)))
NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)


End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
Dim daño As Integer
Dim ANpc As npc, DNpc As npc
ANpc = Npclist(Atacante)

daño = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHIT)
Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - daño

If Npclist(Victima).Stats.MinHP < 1 Then
        
        If Npclist(Atacante).flags.AttackedBy <> "" Then
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
            Npclist(Atacante).Hostile = Npclist(Atacante).flags.OldHostil
        Else
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
        End If
        
        Call FollowAmo(Atacante)
        
        Call MuereNpc(Victima, Npclist(Atacante).MaestroUser)
End If

End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)

' El npc puede atacar ???
If Npclist(Atacante).CanAttack = 1 Then
       Npclist(Atacante).CanAttack = 0
        If cambiarMOvimiento Then
            Npclist(Victima).TargetNPC = Atacante
            Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
        End If
Else
    Exit Sub
End If

If Npclist(Atacante).flags.Snd1 > 0 Then Call SendData(SendTarget.ToNPCArea, Atacante, Npclist(Atacante).pos.Map, "TW" & Npclist(Atacante).flags.Snd1)

If NpcImpactoNpc(Atacante, Victima) Then
    
    If Npclist(Victima).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, Victima, Npclist(Victima).pos.Map, "TW" & Npclist(Victima).flags.Snd2)
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, Npclist(Victima).pos.Map, "TW" & SND_IMPACTO2)
    End If

    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, Npclist(Atacante).pos.Map, "TW" & SND_IMPACTO)
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, Npclist(Victima).pos.Map, "TW" & SND_IMPACTO)
    End If
    Call NpcDañoNpc(Atacante, Victima)
    
Else
    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, Npclist(Atacante).pos.Map, "TW" & SND_SWING)
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, Npclist(Victima).pos.Map, "TW" & SND_SWING)
    End If
End If

End Sub

Public Sub UsuarioAtacaNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer)
 
If Distancia(UserList(userindex).pos, Npclist(NpcIndex).pos) > MAXDISTANCIAARCO Then
   Call SendData(SendTarget.toindex, userindex, 0, "|-Estás muy lejos para disparar." & FONTTYPE_FIGHT)
   Exit Sub
End If

If Npclist(NpcIndex).NPCtype = ReyCastillo Then
        If (Npclist(NpcIndex).pos.Map = MapCastilloN Or Npclist(NpcIndex).pos.Map = MapCastilloS Or Npclist(NpcIndex).pos.Map = MapCastilloE Or Npclist(NpcIndex).pos.Map = MapCastilloO) Then
            Dim castiact As String
            If Npclist(NpcIndex).pos.Map = MapCastilloN Then castiact = CastilloNorte
            If Npclist(NpcIndex).pos.Map = MapCastilloS Then castiact = CastilloSur
            If Npclist(NpcIndex).pos.Map = MapCastilloE Then castiact = CastilloEste
            If Npclist(NpcIndex).pos.Map = MapCastilloO Then castiact = CastilloOeste
 
If Not UserList(userindex).GuildIndex <> 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar al rey del castillo por que no perteneses a ningun clan!!" & FONTTYPE_FIGHT)
             Exit Sub
            End If
            If Guilds(UserList(userindex).GuildIndex).GuildName = castiact Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar al rey de tu castillo " & FONTTYPE_FIGHT)
                Exit Sub
            End If
        End If
    End If
 
If UserList(userindex).flags.Seguro And Npclist(NpcIndex).MaestroUser <> 0 Then
    If Not Criminal(Npclist(NpcIndex).MaestroUser) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Debes sacar el seguro antes de poder atacar una mascota de un ciudadano." & FONTTYPE_WARNING)
        Exit Sub
    End If
End If
 
If UserList(userindex).Faccion.ArmadaReal = 1 And Npclist(NpcIndex).MaestroUser <> 0 Then
    If Not Criminal(Npclist(NpcIndex).MaestroUser) Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Los soldados del Ejercito Real tienen prohibido atacar ciudadanos y sus macotas." & FONTTYPE_WARNING)
        Exit Sub
    End If
End If
 
If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal And UserList(userindex).flags.Seguro Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Debes quitar el seguro para atacar guardias." & FONTTYPE_FIGHT)
    Exit Sub
End If
 
 
Call NpcAtacado(NpcIndex, userindex)
 
If UserImpactoNpc(userindex, NpcIndex) Then
   
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
        Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "AP" & UserList(userindex).Char.CharIndex)
    Else
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_IMPACTO2)
        Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "AP" & UserList(userindex).Char.CharIndex)
    End If
   
    Call UserDañoNpc(userindex, NpcIndex)
   
Else
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_SWING)
    Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "AP" & UserList(userindex).Char.CharIndex)
    Call SendData(SendTarget.toindex, userindex, 0, "U1")
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & &HFFFF& & "°¡Fallo!" & "°" & str(Npclist(NpcIndex).Char.CharIndex))
End If
 
End Sub
Public Sub UsuarioAtaca(ByVal userindex As Integer)

'If UserList(UserIndex).flags.PuedeAtacar = 1 Then
If IntervaloPermiteAtacar(userindex) Then
    
    'Quitamos stamina
    If UserList(userindex).Stats.MinSta >= 10 Then
        Call QuitarSta(userindex, RandomNumber(1, 10))
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||Estas muy cansado para luchar." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'UserList(UserIndex).flags.PuedeAtacar = 0
    
    Dim AttackPos As WorldPos
    AttackPos = UserList(userindex).pos
    Call HeadtoPos(UserList(userindex).Char.Heading, AttackPos)
    
    'Exit if not legal
    If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_SWING)
        Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "AP" & UserList(userindex).Char.CharIndex)
        Exit Sub
    End If
    
    Dim index As Integer
    index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).userindex
        
    'Look for user
    If index > 0 Then
    If UserList(userindex).flags.GemaActivada = "Violeta" Then
        UserEnvenena userindex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).userindex
    End If
        Call UsuarioAtacaUsuario(userindex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).userindex)
        Call SendUserStatsBox(userindex)
        Exit Sub
    End If
    
    'Look for NPC
    If MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex > 0 Then
    
        If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).Attackable Then
            
            If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).MaestroUser > 0 And _
               MapInfo(Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).pos.Map).Pk = False Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No podés atacar mascotas en zonas seguras" & FONTTYPE_FIGHT)
                    Exit Sub
            End If

            Call UsuarioAtacaNpc(userindex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex)
            
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||No podés atacar a este NPC" & FONTTYPE_FIGHT)
        End If
        
        Call SendUserStatsBox(userindex)
        
        Exit Sub
    End If
    
    
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_SWING)
    Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "AP" & UserList(userindex).Char.CharIndex)
    Call SendUserStatsBox(userindex)
End If

If UserList(userindex).Counters.Trabajando Then _
    UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando - 1
    
If UserList(userindex).Counters.Ocultando Then _
    UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando - 1

End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean

Dim ProbRechazo As Long
Dim Rechazo As Boolean
Dim ProbExito As Long
Dim PoderAtaque As Long
Dim UserPoderEvasion As Long
Dim UserPoderEvasionEscudo As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim SkillTacticas As Long
Dim SkillDefensa As Long

SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)

Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
If Arma > 0 Then
    proyectil = ObjData(Arma).proyectil = 1
Else
    proyectil = False
End If

'Calculamos el poder de evasion...
UserPoderEvasion = PoderEvasion(VictimaIndex)

If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
   UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
   UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
Else
    UserPoderEvasionEscudo = 0
End If

'Esta usando un arma ???
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
    Else
        PoderAtaque = PoderAtaqueArma(AtacanteIndex)
    End If
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
   
Else
    PoderAtaque = PoderAtaqueWresterling(AtacanteIndex)
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
    
End If
UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
    
    'Fallo ???
    If UsuarioImpacto = False Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo = True Then
      'Se rechazo el ataque con el escudo
              Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "TW" & SND_ESCUDO)
              Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "EM" & UserList(VictimaIndex).Char.CharIndex)
              Call SendData(SendTarget.toindex, AtacanteIndex, 0, "8")
              Call SendData(SendTarget.toindex, VictimaIndex, 0, "7")
              Call SubirSkill(VictimaIndex, Defensa)
      End If
    End If
End If
    
If UsuarioImpacto Then
   If Arma > 0 Then
           If Not proyectil Then
                  Call SubirSkill(AtacanteIndex, Armas)
           Else
                  Call SubirSkill(AtacanteIndex, Proyectiles)
           End If
   Else
        Call SubirSkill(AtacanteIndex, Wresterling)
   End If
End If

End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub

If Distancia(UserList(AtacanteIndex).pos, UserList(VictimaIndex).pos) > MAXDISTANCIAARCO Then
   Call SendData(SendTarget.toindex, AtacanteIndex, 0, "||Estás muy lejos para disparar." & FONTTYPE_FIGHT)
   Exit Sub
End If


Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
    Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "TW" & SND_IMPACTO)
    
    If UserList(VictimaIndex).flags.Navegando = 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, UserList(VictimaIndex).pos.Map, "XFC" & UserList(VictimaIndex).Char.CharIndex & "," & FXSANGRE & "," & 0)
    If UserList(VictimaIndex).flags.Montando = 0 Then Call SendData(ToPCArea, VictimaIndex, UserList(VictimaIndex).pos.Map, "CFX" & UserList(VictimaIndex).Char.CharIndex & "," & FXSANGRE & "," & 0)
    
    Call UserDañoUser(AtacanteIndex, VictimaIndex)
Else
    Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "TW" & SND_SWING)
    Call SendData(SendTarget.toindex, AtacanteIndex, 0, "U1")
    Call SendData(SendTarget.toindex, VictimaIndex, 0, "U3" & UserList(AtacanteIndex).name)
    Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "||" & &HFFFF& & "°¡Fallo!" & "°" & str(UserList(VictimaIndex).Char.CharIndex))
End If

Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "AP" & UserList(AtacanteIndex).Char.CharIndex)
If UCase$(UserList(AtacanteIndex).Clase) = "LADRON" Then Call Desarmar(AtacanteIndex, VictimaIndex)

End Sub

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim daño As Long, antdaño As Integer
Dim Lugar As Integer, absorbido As Long
Dim defbarco As Integer, defMontura As Integer

Dim Obj As ObjData

daño = CalcularDaño(AtacanteIndex)
antdaño = daño

Call UserEnvenena(AtacanteIndex, VictimaIndex)
Call UserParaliza(AtacanteIndex, VictimaIndex)

If UserList(AtacanteIndex).flags.Navegando = 1 Then
     Obj = ObjData(UserList(AtacanteIndex).Invent.BarcoObjIndex)
     daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
End If

If UserList(VictimaIndex).flags.Navegando = 1 Then
     Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
     defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

If UserList(AtacanteIndex).flags.Montando = 1 Then
     Obj = ObjData(UserList(AtacanteIndex).Invent.MonturaObjIndex)
     daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
End If
 
If UserList(VictimaIndex).flags.Montando = 1 Then
     Obj = ObjData(UserList(VictimaIndex).Invent.MonturaObjIndex)
     defMontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

Dim Resist As Byte
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    Resist = ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Refuerzo
End If

Lugar = RandomNumber(1, 6)

Select Case Lugar
  
  Case PartesCuerpo.bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco - Resist
           daño = daño - absorbido
           If daño < 0 Then daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
           Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco - Resist
           daño = daño - absorbido
           If daño < 0 Then daño = 1
        End If
End Select

If UserList(VictimaIndex).flags.GemaActivada = "Verde" Then
daño = daño - (daño * 10 / 100 + RandomNumber(0, 4))
 
End If

Call SendData(SendTarget.toindex, AtacanteIndex, 0, "N5" & Lugar & "," & daño & "," & UserList(VictimaIndex).name)
Call SendData(SendTarget.toindex, VictimaIndex, 0, "N4" & Lugar & "," & daño & "," & UserList(AtacanteIndex).name)

UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - daño
        
        'Trata de apuñalar por la espalda al enemigo
        If PuedeApuñalar(AtacanteIndex) Then
                Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, daño)
                Call SubirSkill(AtacanteIndex, Apuñalar)
        End If


If UserList(VictimaIndex).Stats.MinHP <= 0 Then
    
    Call ContarMuerte(VictimaIndex, AtacanteIndex)
    
    ' Para que las mascotas no sigan intentando luchar y
    ' comiencen a seguir al amo
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(AtacanteIndex).MascotasIndex(j) > 0 Then
            If Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = VictimaIndex Then Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = 0
            Call FollowAmo(UserList(AtacanteIndex).MascotasIndex(j))
        End If
    Next j
    
    Call ActStats(VictimaIndex, AtacanteIndex)
Else
    'Está vivo - Actualizamos el HP
    Call SendData(SendTarget.toindex, VictimaIndex, 0, "ASH" & UserList(VictimaIndex).Stats.MinHP)
End If

'Controla el nivel del usuario
Call CheckUserLevel(AtacanteIndex)

Call SendData(SendTarget.ToPCArea, AtacanteIndex, UserList(AtacanteIndex).pos.Map, "||" & &HFFFF& & "°" & "- " & daño & "" & "°" & str(UserList(VictimaIndex).Char.CharIndex))

End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
    If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    If Not Criminal(VictimIndex) And UserList(AttackerIndex).pos.Map <> 31 And UserList(AttackerIndex).pos.Map <> 32 And UserList(AttackerIndex).pos.Map <> 33 And UserList(AttackerIndex).pos.Map <> 34 Then
        UserList(AttackerIndex).Reputacion.BandidoRep = UserList(AttackerIndex).Reputacion.BandidoRep + vlASALTO
        UserList(AttackerIndex).Reputacion.BurguesRep = 0
         UserList(AttackerIndex).Reputacion.NobleRep = 0
         UserList(AttackerIndex).Reputacion.PlebeRep = 0
        If UserList(AttackerIndex).Reputacion.BandidoRep > MAXREP Then _
            UserList(AttackerIndex).Reputacion.BandidoRep = MAXREP
    Else
        UserList(AttackerIndex).Reputacion.NobleRep = UserList(AttackerIndex).Reputacion.NobleRep + vlNoble
        If UserList(AttackerIndex).Reputacion.NobleRep > MAXREP Then _
            UserList(AttackerIndex).Reputacion.NobleRep = MAXREP
    End If
    
    Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
       If UserList(AttackerIndex).flags.EstaDueleando = True And UserList(VictimIndex).flags.EstaDueleando = True Then
    Exit Sub
    End If
End Sub

Sub AllMascotasAtacanUser(ByVal Victim As Integer, ByVal Maestro As Integer)
'Reaccion de las mascotas
Dim iCount As Integer

For iCount = 1 To MAXMASCOTAS
    If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(Victim).name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
    End If
Next iCount

End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

Dim T As eTrigger6

If UserList(VictimIndex).flags.Muerto = 1 Then
    SendData SendTarget.toindex, AttackerIndex, 0, "||No podes atacar a un espiritu" & FONTTYPE_INFO
    PuedeAtacar = False
    Exit Function
End If

Dim klan As String
If UserList(AttackerIndex).flags.SeguroClan = True Then
If UserList(VictimIndex).GuildIndex > 0 Then
If UserList(AttackerIndex).GuildIndex > 0 Then
    klan = Guilds(UserList(AttackerIndex).GuildIndex).GuildName
   
    If Guilds(UserList(VictimIndex).GuildIndex).GuildName = Guilds(UserList(AttackerIndex).GuildIndex).GuildName And klan <> "" Then
        Call SendData(SendTarget.toindex, AttackerIndex, 0, "||No puedes atacar a tu propio Clan con el seguro activado, escribe /SEGCLAN para desactivarlo." & FONTTYPE_FIGHT)
        Exit Function
    End If
    End If
    End If
End If

T = TriggerZonaPelea(AttackerIndex, VictimIndex)

If T = TRIGGER6_PERMITE Then
    PuedeAtacar = True
    Exit Function
ElseIf T = TRIGGER6_PROHIBE Then
    PuedeAtacar = False
    Exit Function
End If


If MapInfo(UserList(VictimIndex).pos.Map).Pk = False Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||Esta es una zona segura, aqui no podes atacar otros usuarios." & FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

If MapData(UserList(VictimIndex).pos.Map, UserList(VictimIndex).pos.X, UserList(VictimIndex).pos.Y).trigger = eTrigger.ZONASEGURA Or _
    MapData(UserList(AttackerIndex).pos.Map, UserList(AttackerIndex).pos.X, UserList(AttackerIndex).pos.Y).trigger = eTrigger.ZONASEGURA Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||No podes pelear aqui." & FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

If Ciudadano(VictimIndex) And UserList(AttackerIndex).Faccion.ArmadaReal = 1 And UserList(AttackerIndex).pos.Map <> MapCastilloS And UserList(AttackerIndex).pos.Map <> MapCastilloN And UserList(AttackerIndex).pos.Map <> MapCastilloE And UserList(AttackerIndex).pos.Map <> MapCastilloO Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||Los soldados del Ejercito Real tienen prohibido atacar ciudadanos." & FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

'[Barrin 17-12-03]
If (UserList(VictimIndex).Faccion.FuerzasCaos = 1) And (UserList(AttackerIndex).Faccion.FuerzasCaos = 1) And UserList(AttackerIndex).pos.Map <> MapCastilloS And UserList(AttackerIndex).pos.Map <> MapCastilloN And UserList(AttackerIndex).pos.Map <> MapCastilloE And UserList(AttackerIndex).pos.Map <> MapCastilloO Then
    Call SendData(SendTarget.toindex, AttackerIndex, 0, "||Los miembros de la legión oscura tienen prohibido atacarse entre sí." & FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If
'[/Barrin]

'Se asegura que la victima no es un GM
If UserList(VictimIndex).flags.Privilegios >= PlayerType.SemiDios Then
    SendData SendTarget.toindex, AttackerIndex, 0, "||¡¡No podes atacar a los administradores del juego!! " & FONTTYPE_WARNING
    PuedeAtacar = False
    Exit Function
End If

If UserList(AttackerIndex).flags.Muerto = 1 Then
    SendData SendTarget.toindex, AttackerIndex, 0, "||No podes atacar porque estas muerto" & FONTTYPE_INFO
    PuedeAtacar = False
    Exit Function
End If

If Ciudadano(VictimIndex) And Ciudadano(AttackerIndex) And UserList(AttackerIndex).pos.Map <> MapCastilloS And UserList(AttackerIndex).pos.Map <> MapCastilloN And UserList(AttackerIndex).pos.Map <> MapCastilloE And UserList(AttackerIndex).pos.Map <> MapCastilloO Then
                Call SendData(SendTarget.toindex, AttackerIndex, 0, "||No podes atacar ciudadanos, si deseas hacerlo escribe /RENUNCIAR." & FONTTYPE_FIGHT)
                Exit Function
        End If
   

PuedeAtacar = True

End Function


Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser > 0 Then
    If Not Criminal(AttackerIndex) And Not Criminal(Npclist(NpcIndex).MaestroUser) Then
        If UserList(AttackerIndex).flags.Seguro Then
            Call SendData(SendTarget.toindex, AttackerIndex, 0, "||Para atacar mascotas de ciudadanos debes quitarte el seguro" & FONTTYPE_FIGHT)
            PuedeAtacarNPC = False
            Exit Function
        End If
    End If
End If

If UserList(AttackerIndex).flags.Muerto = 1 Then
    SendData SendTarget.toindex, AttackerIndex, 0, "||No podes atacar porque estas muerto" & FONTTYPE_INFO
    PuedeAtacarNPC = False
    Exit Function
End If


PuedeAtacarNPC = True

End Function


'[KEVIN]
'
'[Alejo]
'Modifique un poco el sistema de exp por golpe, ahora
'son 2/3 de la exp mientras esta vivo, el resto se
'obtiene al matarlo.
'Ahora además
Sub CalcularDarExp(ByVal userindex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)

Dim ExpSinMorir As Long
Dim ExpaDar As Long
Dim TotalNpcVida As Long

If ElDaño <= 0 Then ElDaño = 0

'2/3 de la experiencia se dan cuando se le golpea, el resto
'se obtiene al matarlo
ExpSinMorir = (2 * Npclist(NpcIndex).GiveEXP) / 3

TotalNpcVida = Npclist(NpcIndex).Stats.MaxHP
If TotalNpcVida <= 0 Then Exit Sub

If ElDaño > Npclist(NpcIndex).Stats.MinHP Then ElDaño = Npclist(NpcIndex).Stats.MinHP

'totalnpcvida _____ ExpSinMorir
'daño         _____ (daño * ExpSinMorir) / totalNpcVida

ExpaDar = CLng((ElDaño) * (ExpSinMorir / TotalNpcVida))
If ExpaDar <= 0 Then Exit Sub

If UserList(userindex).flags.GemaActivada = "Plateada" Then
 
    ExpaDar = ExpaDar + (ExpaDar / 2)
   
End If

If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
    ExpaDar = Npclist(NpcIndex).flags.ExpCount
    Npclist(NpcIndex).flags.ExpCount = 0
Else
    Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar
End If

If ExpaDar > 0 Then
    If UserList(userindex).PartyIndex > 0 Then
        Call mdParty.ObtenerExito(userindex, ExpaDar, Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y)
    Else
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpaDar
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
        Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado " & ExpaDar & " puntos de experiencia." & FONTTYPE_AMARILLON)
    End If
    
    Call CheckUserLevel(userindex)
End If

'[/KEVIN]
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6

If Origen > 0 And Destino > 0 And Origen <= UBound(UserList) And Destino <= UBound(UserList) Then
    If MapData(UserList(Origen).pos.Map, UserList(Origen).pos.X, UserList(Origen).pos.Y).trigger = eTrigger.ZONAPELEA Or _
        MapData(UserList(Destino).pos.Map, UserList(Destino).pos.X, UserList(Destino).pos.Y).trigger = eTrigger.ZONAPELEA Then
        If (MapData(UserList(Origen).pos.Map, UserList(Origen).pos.X, UserList(Origen).pos.Y).trigger = MapData(UserList(Destino).pos.Map, UserList(Destino).pos.X, UserList(Destino).pos.Y).trigger) Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If
Else
    TriggerZonaPelea = TRIGGER6_AUSENTE
End If

End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim ArmaObjInd As Integer, ObjInd As Integer
Dim num As Long

ArmaObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
ObjInd = 0

If ArmaObjInd > 0 Then
    If ObjData(ArmaObjInd).proyectil = 0 Then
        ObjInd = ArmaObjInd
    Else
        ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
    End If
    
    If ObjInd > 0 Then
        If (ObjData(ObjInd).Envenena = 1) Then
            num = RandomNumber(1, 100)
            
            If num < 60 Then
                UserList(VictimaIndex).flags.Envenenado = 1
                Call SendData(SendTarget.toindex, VictimaIndex, 0, "||" & UserList(AtacanteIndex).name & " te ha envenenado!!" & FONTTYPE_FIGHT)
                Call SendData(SendTarget.toindex, AtacanteIndex, 0, "||Has envenenado a " & UserList(VictimaIndex).name & "!!" & FONTTYPE_FIGHT)
            End If
        End If
    End If
End If

End Sub

Sub UserParaliza(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim ArmaObjInd As Integer, ObjInd As Integer
Dim num As Long

ArmaObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
ObjInd = 0

If ArmaObjInd > 0 Then
If ObjData(ArmaObjInd).proyectil = 0 Then
ObjInd = ArmaObjInd
Else
ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
End If

If ObjInd > 0 Then
If (ObjData(ObjInd).Paraliza = 1) Then
num = RandomNumber(1, 100)

If num < 60 Then
UserList(VictimaIndex).flags.Paralizado = 1
UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado
Call SendData(toindex, VictimaIndex, 0, "PARADOK")
Call SendData(toindex, VictimaIndex, 0, "||" & UserList(AtacanteIndex).name & " te ha paralizado!!" & FONTTYPE_FIGHT)
Call SendData(toindex, AtacanteIndex, 0, "||Has paralizado a " & UserList(VictimaIndex).name & "!!" & FONTTYPE_FIGHT)
End If
End If
End If
End If

End Sub
