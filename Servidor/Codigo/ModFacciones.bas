Attribute VB_Name = "ModFacciones"
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


Option Explicit
                                     'Imperial
'Bajos
Public GuerreroBajosArmadura1 As Integer
Public MagoBajosArmadura1 As Integer
Public PaladinBajosArmadura1 As Integer
Public ClerigoBajosArmadura1 As Integer
Public BardoBajosArmadura1 As Integer
Public AsesinoBajosArmadura1 As Integer
Public DruidaBajosArmadura1 As Integer
Public CazadorBajosArmadura1 As Integer
Public OtrasBajosArmadura1 As Integer

Public GuerreroBajosArmadura2 As Integer
Public MagoBajosArmadura2 As Integer
Public PaladinBajosArmadura2 As Integer
Public ClerigoBajosArmadura2 As Integer
Public BardoBajosArmadura2 As Integer
Public AsesinoBajosArmadura2 As Integer
Public DruidaBajosArmadura2 As Integer
Public CazadorBajosArmadura2 As Integer
Public OtrasBajosArmadura2 As Integer

Public GuerreroBajosArmadura3 As Integer
Public MagoBajosArmadura3 As Integer
Public PaladinBajosArmadura3 As Integer
Public ClerigoBajosArmadura3 As Integer
Public BardoBajosArmadura3 As Integer
Public AsesinoBajosArmadura3 As Integer
Public DruidaBajosArmadura3 As Integer
Public CazadorBajosArmadura3 As Integer
Public OtrasBajosArmadura3 As Integer

Public GuerreroBajosArmadura4 As Integer
Public MagoBajosArmadura4 As Integer
Public PaladinBajosArmadura4 As Integer
Public ClerigoBajosArmadura4 As Integer
Public BardoBajosArmadura4 As Integer
Public AsesinoBajosArmadura4 As Integer
Public DruidaBajosArmadura4 As Integer
Public CazadorBajosArmadura4 As Integer
Public OtrasBajosArmadura4 As Integer

'Altos
Public GuerreroAltosArmadura1 As Integer
Public MagoAltosArmadura1 As Integer
Public PaladinAltosArmadura1 As Integer
Public ClerigoAltosArmadura1 As Integer
Public BardoAltosArmadura1 As Integer
Public AsesinoAltosArmadura1 As Integer
Public DruidaAltosArmadura1 As Integer
Public CazadorAltosArmadura1 As Integer
Public OtrasAltosArmadura1 As Integer

Public GuerreroAltosArmadura2 As Integer
Public MagoAltosArmadura2 As Integer
Public PaladinAltosArmadura2 As Integer
Public ClerigoAltosArmadura2 As Integer
Public BardoAltosArmadura2 As Integer
Public AsesinoAltosArmadura2 As Integer
Public DruidaAltosArmadura2 As Integer
Public CazadorAltosArmadura2 As Integer
Public OtrasAltosArmadura2 As Integer

Public GuerreroAltosArmadura3 As Integer
Public MagoAltosArmadura3 As Integer
Public PaladinAltosArmadura3 As Integer
Public ClerigoAltosArmadura3 As Integer
Public BardoAltosArmadura3 As Integer
Public AsesinoAltosArmadura3 As Integer
Public DruidaAltosArmadura3 As Integer
Public CazadorAltosArmadura3 As Integer
Public OtrasAltosArmadura3 As Integer

Public GuerreroAltosArmadura4 As Integer
Public MagoAltosArmadura4 As Integer
Public PaladinAltosArmadura4 As Integer
Public ClerigoAltosArmadura4 As Integer
Public BardoAltosArmadura4 As Integer
Public AsesinoAltosArmadura4 As Integer
Public DruidaAltosArmadura4 As Integer
Public CazadorAltosArmadura4 As Integer
Public OtrasAltosArmadura4 As Integer

                                     'Caotico
'Bajos
Public CaosGuerreroBajosArmadura1 As Integer
Public CaosMagoBajosArmadura1 As Integer
Public CaosPaladinBajosArmadura1 As Integer
Public CaosClerigoBajosArmadura1 As Integer
Public CaosBardoBajosArmadura1 As Integer
Public CaosAsesinoBajosArmadura1 As Integer
Public CaosDruidaBajosArmadura1 As Integer
Public CaosCazadorBajosArmadura1 As Integer
Public CaosOtrasBajosArmadura1 As Integer

Public CaosGuerreroBajosArmadura2 As Integer
Public CaosMagoBajosArmadura2 As Integer
Public CaosPaladinBajosArmadura2 As Integer
Public CaosClerigoBajosArmadura2 As Integer
Public CaosBardoBajosArmadura2 As Integer
Public CaosAsesinoBajosArmadura2 As Integer
Public CaosDruidaBajosArmadura2 As Integer
Public CaosCazadorBajosArmadura2 As Integer
Public CaosOtrasBajosArmadura2 As Integer

Public CaosGuerreroBajosArmadura3 As Integer
Public CaosMagoBajosArmadura3 As Integer
Public CaosPaladinBajosArmadura3 As Integer
Public CaosClerigoBajosArmadura3 As Integer
Public CaosBardoBajosArmadura3 As Integer
Public CaosAsesinoBajosArmadura3 As Integer
Public CaosDruidaBajosArmadura3 As Integer
Public CaosCazadorBajosArmadura3 As Integer
Public CaosOtrasBajosArmadura3 As Integer

Public CaosGuerreroBajosArmadura4 As Integer
Public CaosMagoBajosArmadura4 As Integer
Public CaosPaladinBajosArmadura4 As Integer
Public CaosClerigoBajosArmadura4 As Integer
Public CaosBardoBajosArmadura4 As Integer
Public CaosAsesinoBajosArmadura4 As Integer
Public CaosDruidaBajosArmadura4 As Integer
Public CaosCazadorBajosArmadura4 As Integer
Public CaosOtrasBajosArmadura4 As Integer

'Altos
Public CaosGuerreroAltosArmadura1 As Integer
Public CaosMagoAltosArmadura1 As Integer
Public CaosPaladinAltosArmadura1 As Integer
Public CaosClerigoAltosArmadura1 As Integer
Public CaosBardoAltosArmadura1 As Integer
Public CaosAsesinoAltosArmadura1 As Integer
Public CaosDruidaAltosArmadura1 As Integer
Public CaosCazadorAltosArmadura1 As Integer
Public CaosOtrasAltosArmadura1 As Integer

Public CaosGuerreroAltosArmadura2 As Integer
Public CaosMagoAltosArmadura2 As Integer
Public CaosPaladinAltosArmadura2 As Integer
Public CaosClerigoAltosArmadura2 As Integer
Public CaosBardoAltosArmadura2 As Integer
Public CaosAsesinoAltosArmadura2 As Integer
Public CaosDruidaAltosArmadura2 As Integer
Public CaosCazadorAltosArmadura2 As Integer
Public CaosOtrasAltosArmadura2 As Integer

Public CaosGuerreroAltosArmadura3 As Integer
Public CaosMagoAltosArmadura3 As Integer
Public CaosPaladinAltosArmadura3 As Integer
Public CaosClerigoAltosArmadura3 As Integer
Public CaosBardoAltosArmadura3 As Integer
Public CaosAsesinoAltosArmadura3 As Integer
Public CaosDruidaAltosArmadura3 As Integer
Public CaosCazadorAltosArmadura3 As Integer
Public CaosOtrasAltosArmadura3 As Integer

Public CaosGuerreroAltosArmadura4 As Integer
Public CaosMagoAltosArmadura4 As Integer
Public CaosPaladinAltosArmadura4 As Integer
Public CaosClerigoAltosArmadura4 As Integer
Public CaosBardoAltosArmadura4 As Integer
Public CaosAsesinoAltosArmadura4 As Integer
Public CaosDruidaAltosArmadura4 As Integer
Public CaosCazadorAltosArmadura4 As Integer
Public CaosOtrasAltosArmadura4

Public Const ExpAlUnirse As Long = 50000
Public Const ExpX100 As Integer = 5000


Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya perteneces a las tropas reales!!! Ve a combatir criminales!!!" & FONTTYPE_UDP)
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Maldito insolente!!! vete de aqui seguidor de las sombras!!!" & FONTTYPE_UDP)
    Exit Sub
End If

If Not Ciudadano(UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo se permiten ciudadanos en el ejercito real!!!" & FONTTYPE_UDP)
    Exit Sub
End If

'If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Then
 '   Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
  '  Exit Sub
'End If

If UserList(UserIndex).Faccion.CriminalesMatados <= 49 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para enlistarte en las tropas reales tienes que a ver matado 50 criminales." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
Exit Sub
End If

If UserList(UserIndex).Faccion.Reenlistadas = 1 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo puedes enlistarte una vez." & FONTTYPE_UDP)
Exit Sub
End If

UserList(UserIndex).Faccion.ArmadaReal = 1
UserList(UserIndex).Faccion.Reenlistadas = 1
UserList(UserIndex).Faccion.RecompensasReal = 0

Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡¡Bienvenido a al Ejercito Imperial escribe /recompensa para recibir tu armadura!!!" & FONTTYPE_UDP)

If UserList(UserIndex).Faccion.RecibioExpInicialReal = 0 Then
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpAlUnirse
    If UserList(UserIndex).Stats.Exp > MAXEXP Then _
        UserList(UserIndex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialReal = 1
    Call CheckUserLevel(UserIndex)
End If

Call LogEjercitoReal(UserList(UserIndex).name)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
Dim MiObj As Obj
    MiObj.Amount = 1
If UserList(UserIndex).Faccion.CriminalesMatados \ 50 = _
   UserList(UserIndex).Faccion.RecompensasReal Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata 50 crinales mas para recibir la proxima!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100
    If UserList(UserIndex).Stats.Exp > MAXEXP Then _
        UserList(UserIndex).Stats.Exp = MAXEXP
        
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.RecompensasReal + 1
    
    Select Case UserList(UserIndex).Faccion.RecompensasReal
    
    Case 1

    If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = GuerreroBajosArmadura1
            Case "MAGO"
            MiObj.ObjIndex = MagoBajosArmadura1
            Case "PALADIN"
            MiObj.ObjIndex = PaladinBajosArmadura1
            Case "CLERIGO"
            MiObj.ObjIndex = ClerigoBajosArmadura1
            Case "BARDO"
            MiObj.ObjIndex = BardoBajosArmadura1
            Case "ASESINO"
            MiObj.ObjIndex = AsesinoBajosArmadura1
            Case "DRUIDA"
            MiObj.ObjIndex = DruidaBajosArmadura1
            Case "CAZADOR"
            MiObj.ObjIndex = CazadorBajosArmadura1
            Case Else
            MiObj.ObjIndex = OtrasBajosArmadura1
        End Select
     Else
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = GuerreroAltosArmadura1
            Case "MAGO"
            MiObj.ObjIndex = MagoAltosArmadura1
            Case "PALADIN"
            MiObj.ObjIndex = PaladinAltosArmadura1
            Case "CLERIGO"
            MiObj.ObjIndex = ClerigoAltosArmadura1
            Case "BARDO"
            MiObj.ObjIndex = BardoAltosArmadura1
            Case "ASESINO"
            MiObj.ObjIndex = AsesinoAltosArmadura1
            Case "DRUIDA"
            MiObj.ObjIndex = DruidaAltosArmadura1
            Case "CAZADOR"
            MiObj.ObjIndex = CazadorAltosArmadura1
            Case Else
            MiObj.ObjIndex = OtrasAltosArmadura1
        End Select
        End If
     
     Case 2
     
         If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = GuerreroBajosArmadura2
            Case "MAGO"
            MiObj.ObjIndex = MagoBajosArmadura2
            Case "PALADIN"
            MiObj.ObjIndex = PaladinBajosArmadura2
            Case "CLERIGO"
            MiObj.ObjIndex = ClerigoBajosArmadura2
            Case "BARDO"
            MiObj.ObjIndex = BardoBajosArmadura2
            Case "ASESINO"
            MiObj.ObjIndex = AsesinoBajosArmadura2
            Case "DRUIDA"
            MiObj.ObjIndex = DruidaBajosArmadura2
            Case "CAZADOR"
            MiObj.ObjIndex = CazadorBajosArmadura2
            Case Else
            MiObj.ObjIndex = OtrasBajosArmadura2
        End Select
     Else
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = GuerreroAltosArmadura2
            Case "MAGO"
            MiObj.ObjIndex = MagoAltosArmadura2
            Case "PALADIN"
            MiObj.ObjIndex = PaladinAltosArmadura2
            Case "CLERIGO"
            MiObj.ObjIndex = ClerigoAltosArmadura2
            Case "BARDO"
            MiObj.ObjIndex = BardoAltosArmadura2
            Case "ASESINO"
            MiObj.ObjIndex = AsesinoAltosArmadura2
            Case "DRUIDA"
            MiObj.ObjIndex = DruidaAltosArmadura2
            Case "CAZADOR"
            MiObj.ObjIndex = CazadorAltosArmadura2
            Case Else
            MiObj.ObjIndex = OtrasAltosArmadura2
        End Select
        End If
    
     Case 3
     
         If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = GuerreroBajosArmadura3
            Case "MAGO"
            MiObj.ObjIndex = MagoBajosArmadura3
            Case "PALADIN"
            MiObj.ObjIndex = PaladinBajosArmadura3
            Case "CLERIGO"
            MiObj.ObjIndex = ClerigoBajosArmadura3
            Case "BARDO"
            MiObj.ObjIndex = BardoBajosArmadura3
            Case "ASESINO"
            MiObj.ObjIndex = AsesinoBajosArmadura3
            Case "DRUIDA"
            MiObj.ObjIndex = DruidaBajosArmadura3
            Case "CAZADOR"
            MiObj.ObjIndex = CazadorBajosArmadura3
            Case Else
            MiObj.ObjIndex = OtrasBajosArmadura3
        End Select
     Else
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = GuerreroAltosArmadura3
            Case "MAGO"
            MiObj.ObjIndex = MagoAltosArmadura3
            Case "PALADIN"
            MiObj.ObjIndex = PaladinAltosArmadura3
            Case "CLERIGO"
            MiObj.ObjIndex = ClerigoAltosArmadura3
            Case "BARDO"
            MiObj.ObjIndex = BardoAltosArmadura3
            Case "ASESINO"
            MiObj.ObjIndex = AsesinoAltosArmadura3
            Case "DRUIDA"
            MiObj.ObjIndex = DruidaAltosArmadura3
            Case "CAZADOR"
            MiObj.ObjIndex = CazadorAltosArmadura3
            Case Else
            MiObj.ObjIndex = OtrasAltosArmadura3
        End Select
        End If
        
        Case 4
     
         If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = GuerreroBajosArmadura4
            Case "MAGO"
            MiObj.ObjIndex = MagoBajosArmadura4
            Case "PALADIN"
            MiObj.ObjIndex = PaladinBajosArmadura4
            Case "CLERIGO"
            MiObj.ObjIndex = ClerigoBajosArmadura4
            Case "BARDO"
            MiObj.ObjIndex = BardoBajosArmadura4
            Case "ASESINO"
            MiObj.ObjIndex = AsesinoBajosArmadura4
            Case "DRUIDA"
            MiObj.ObjIndex = DruidaBajosArmadura4
            Case "CAZADOR"
            MiObj.ObjIndex = CazadorBajosArmadura4
            Case Else
            MiObj.ObjIndex = OtrasBajosArmadura4
        End Select
     Else
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = GuerreroAltosArmadura4
            Case "MAGO"
            MiObj.ObjIndex = MagoAltosArmadura4
            Case "PALADIN"
            MiObj.ObjIndex = PaladinAltosArmadura4
            Case "CLERIGO"
            MiObj.ObjIndex = ClerigoAltosArmadura4
            Case "BARDO"
            MiObj.ObjIndex = BardoAltosArmadura4
            Case "ASESINO"
            MiObj.ObjIndex = AsesinoAltosArmadura4
            Case "DRUIDA"
            MiObj.ObjIndex = DruidaAltosArmadura4
            Case "CAZADOR"
            MiObj.ObjIndex = CazadorAltosArmadura4
            Case Else
            MiObj.ObjIndex = OtrasAltosArmadura4
        End Select
        End If
End Select
MiObj.Amount = 1
 
If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
    Call CheckUserLevel(UserIndex)
End If

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)

    UserList(UserIndex).Faccion.ArmadaReal = 0
    UserList(UserIndex).flags.PJerarquia = 0
    UserList(UserIndex).flags.SJerarquia = 0
    UserList(UserIndex).flags.TJerarquia = 0
    UserList(UserIndex).flags.CJerarquia = 0
    'Call PerderItemsFaccionarios(UserIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has sido expulsado de las tropas reales!!!." & FONTTYPE_FIGHT)
    'Desequipamos la armadura real si está equipada
    If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)

    UserList(UserIndex).Faccion.FuerzasCaos = 0
    UserList(UserIndex).flags.PJerarquia = 0
    UserList(UserIndex).flags.SJerarquia = 0
    UserList(UserIndex).flags.TJerarquia = 0
    UserList(UserIndex).flags.CJerarquia = 0
    'Call PerderItemsFaccionarios(UserIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has sido expulsado de la legión oscura!!!." & FONTTYPE_FIGHT)
    'Desequipamos la armadura real si está equipada
    If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
 
Select Case UserList(UserIndex).Faccion.RecompensasReal
    Case 0
        TituloReal = "Aprendiz"
        UserList(UserIndex).flags.PJerarquia = 1
    Case 1
        TituloReal = "Aprendiz"
    Case 2
        TituloReal = "Caballero"
        UserList(UserIndex).flags.PJerarquia = 0
        UserList(UserIndex).flags.SJerarquia = 1
    Case 3
        TituloReal = "Capitán"
        UserList(UserIndex).flags.SJerarquia = 0
        UserList(UserIndex).flags.TJerarquia = 1
    Case Else
        TituloReal = "Campeón de la Luz"
        UserList(UserIndex).flags.TJerarquia = 0
        UserList(UserIndex).flags.CJerarquia = 1
End Select
 
End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)


If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Las sombras reinaran en este mundo, largate de aqui ciudadano.!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.RecibioExpInicialReal = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No permitiré que ningún insecto real ingrese ¡Traidor del Rey!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If Not Criminal(UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ja ja ja tu no eres bienvenido aqui!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CiudadanosMatados < 50 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos 50 ciudadanos, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.Reenlistadas = 1 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo puedes enlistarte una vez." & FONTTYPE_UDP)
Exit Sub
End If

UserList(UserIndex).Faccion.Reenlistadas = 1
UserList(UserIndex).Faccion.FuerzasCaos = 1
UserList(UserIndex).Faccion.RecompensasCaos = 0

Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Bienvenido a al lado oscuro!!!, para recibir tu armadura escribe /recompensa!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))

If UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0 Then
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpAlUnirse
    If UserList(UserIndex).Stats.Exp > MAXEXP Then _
        UserList(UserIndex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialCaos = 1
    Call CheckUserLevel(UserIndex)
End If


Call LogEjercitoCaos(UserList(UserIndex).name)

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
Dim MiObj As Obj
    MiObj.Amount = 1
If UserList(UserIndex).Faccion.CiudadanosMatados \ 50 = _
   UserList(UserIndex).Faccion.RecompensasCaos Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata 50 ciudadanos mas para recibir la proxima!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100
    If UserList(UserIndex).Stats.Exp > MAXEXP Then _
        UserList(UserIndex).Stats.Exp = MAXEXP
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.RecompensasCaos + 1
    
    Select Case UserList(UserIndex).Faccion.RecompensasCaos
    
    Case 1

    If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = CaosGuerreroBajosArmadura1
            Case "MAGO"
            MiObj.ObjIndex = CaosMagoBajosArmadura1
            Case "PALADIN"
            MiObj.ObjIndex = CaosPaladinBajosArmadura1
            Case "CLERIGO"
            MiObj.ObjIndex = CaosClerigoBajosArmadura1
            Case "BARDO"
            MiObj.ObjIndex = CaosBardoBajosArmadura1
            Case "ASESINO"
            MiObj.ObjIndex = CaosAsesinoBajosArmadura1
            Case "DRUIDA"
            MiObj.ObjIndex = CaosDruidaBajosArmadura1
            Case "CAZADOR"
            MiObj.ObjIndex = CaosCazadorBajosArmadura1
            Case Else
            MiObj.ObjIndex = CaosOtrasBajosArmadura1
        End Select
     Else
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = CaosGuerreroAltosArmadura1
            Case "MAGO"
            MiObj.ObjIndex = CaosMagoAltosArmadura1
            Case "PALADIN"
            MiObj.ObjIndex = CaosPaladinAltosArmadura1
            Case "CLERIGO"
            MiObj.ObjIndex = CaosClerigoAltosArmadura1
            Case "BARDO"
            MiObj.ObjIndex = CaosBardoAltosArmadura1
            Case "ASESINO"
            MiObj.ObjIndex = CaosAsesinoAltosArmadura1
            Case "DRUIDA"
            MiObj.ObjIndex = CaosDruidaAltosArmadura1
            Case "CAZADOR"
            MiObj.ObjIndex = CaosCazadorAltosArmadura1
            Case Else
            MiObj.ObjIndex = CaosOtrasAltosArmadura1
        End Select
        End If
     
     Case 2
     
         If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = CaosGuerreroBajosArmadura2
            Case "MAGO"
            MiObj.ObjIndex = CaosMagoBajosArmadura2
            Case "PALADIN"
            MiObj.ObjIndex = CaosPaladinBajosArmadura2
            Case "CLERIGO"
            MiObj.ObjIndex = CaosClerigoBajosArmadura2
            Case "BARDO"
            MiObj.ObjIndex = CaosBardoBajosArmadura2
            Case "ASESINO"
            MiObj.ObjIndex = CaosAsesinoBajosArmadura2
            Case "DRUIDA"
            MiObj.ObjIndex = CaosDruidaBajosArmadura2
            Case "CAZADOR"
            MiObj.ObjIndex = CaosCazadorBajosArmadura2
            Case Else
            MiObj.ObjIndex = CaosOtrasBajosArmadura2
        End Select
     Else
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = CaosGuerreroAltosArmadura2
            Case "MAGO"
            MiObj.ObjIndex = CaosMagoAltosArmadura2
            Case "PALADIN"
            MiObj.ObjIndex = CaosPaladinAltosArmadura2
            Case "CLERIGO"
            MiObj.ObjIndex = CaosClerigoAltosArmadura2
            Case "BARDO"
            MiObj.ObjIndex = CaosBardoAltosArmadura2
            Case "ASESINO"
            MiObj.ObjIndex = CaosAsesinoAltosArmadura2
            Case "DRUIDA"
            MiObj.ObjIndex = CaosDruidaAltosArmadura2
            Case "CAZADOR"
            MiObj.ObjIndex = CaosCazadorAltosArmadura2
            Case Else
            MiObj.ObjIndex = CaosOtrasAltosArmadura2
        End Select
        End If
    
     Case 3
     
         If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = CaosGuerreroBajosArmadura3
            Case "MAGO"
            MiObj.ObjIndex = CaosMagoBajosArmadura3
            Case "PALADIN"
            MiObj.ObjIndex = CaosPaladinBajosArmadura3
            Case "CLERIGO"
            MiObj.ObjIndex = CaosClerigoBajosArmadura3
            Case "BARDO"
            MiObj.ObjIndex = CaosBardoBajosArmadura3
            Case "ASESINO"
            MiObj.ObjIndex = CaosAsesinoBajosArmadura3
            Case "DRUIDA"
            MiObj.ObjIndex = CaosDruidaBajosArmadura3
            Case "CAZADOR"
            MiObj.ObjIndex = CaosCazadorBajosArmadura3
            Case Else
            MiObj.ObjIndex = CaosOtrasBajosArmadura3
        End Select
     Else
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = CaosGuerreroAltosArmadura3
            Case "MAGO"
            MiObj.ObjIndex = CaosMagoAltosArmadura3
            Case "PALADIN"
            MiObj.ObjIndex = CaosPaladinAltosArmadura3
            Case "CLERIGO"
            MiObj.ObjIndex = CaosClerigoAltosArmadura3
            Case "BARDO"
            MiObj.ObjIndex = CaosBardoAltosArmadura3
            Case "ASESINO"
            MiObj.ObjIndex = CaosAsesinoAltosArmadura3
            Case "DRUIDA"
            MiObj.ObjIndex = CaosDruidaAltosArmadura3
            Case "CAZADOR"
            MiObj.ObjIndex = CaosCazadorAltosArmadura3
            Case Else
            MiObj.ObjIndex = CaosOtrasAltosArmadura3
        End Select
        End If
        
        Case 4
     
         If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = CaosGuerreroBajosArmadura4
            Case "MAGO"
            MiObj.ObjIndex = CaosMagoBajosArmadura4
            Case "PALADIN"
            MiObj.ObjIndex = CaosPaladinBajosArmadura4
            Case "CLERIGO"
            MiObj.ObjIndex = CaosClerigoBajosArmadura4
            Case "BARDO"
            MiObj.ObjIndex = CaosBardoBajosArmadura4
            Case "ASESINO"
            MiObj.ObjIndex = CaosAsesinoBajosArmadura4
            Case "DRUIDA"
            MiObj.ObjIndex = CaosDruidaBajosArmadura4
            Case "CAZADOR"
            MiObj.ObjIndex = CaosCazadorBajosArmadura4
            Case Else
            MiObj.ObjIndex = CaosOtrasBajosArmadura4
        End Select
     Else
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "GUERRERO"
            MiObj.ObjIndex = CaosGuerreroAltosArmadura4
            Case "MAGO"
            MiObj.ObjIndex = CaosMagoAltosArmadura4
            Case "PALADIN"
            MiObj.ObjIndex = CaosPaladinAltosArmadura4
            Case "CLERIGO"
            MiObj.ObjIndex = CaosClerigoAltosArmadura4
            Case "BARDO"
            MiObj.ObjIndex = CaosBardoAltosArmadura4
            Case "ASESINO"
            MiObj.ObjIndex = CaosAsesinoAltosArmadura4
            Case "DRUIDA"
            MiObj.ObjIndex = CaosDruidaAltosArmadura4
            Case "CAZADOR"
            MiObj.ObjIndex = CaosCazadorAltosArmadura4
            Case Else
            MiObj.ObjIndex = CaosOtrasAltosArmadura4
        End Select
        End If
End Select
MiObj.Amount = 1
 
If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
    
    Call CheckUserLevel(UserIndex)
End If

End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
Select Case UserList(UserIndex).Faccion.RecompensasCaos
    Case 0
        TituloCaos = "Siervo"
        UserList(UserIndex).flags.PJerarquia = 1
    Case 1
        TituloCaos = "Siervo"
    Case 2
        TituloCaos = "Acólito"
        UserList(UserIndex).flags.PJerarquia = 0
        UserList(UserIndex).flags.SJerarquia = 1
    Case 3
        TituloCaos = "Caballero de la Oscuridad"
        UserList(UserIndex).flags.SJerarquia = 0
        UserList(UserIndex).flags.TJerarquia = 1
    Case Else
        TituloCaos = "Devorador de Almas"
        UserList(UserIndex).flags.TJerarquia = 0
        UserList(UserIndex).flags.CJerarquia = 1
End Select
 
 
End Function
