Attribute VB_Name = "NPCs"
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


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit

Public Const ReyNpcN As Integer = 910 'ACA VA EL NUMERO DE NPC DEL REY By Nait

Sub QuitarMascota(ByVal userindex As Integer, ByVal NpcIndex As Integer)

Dim i As Integer
UserList(userindex).NroMacotas = UserList(userindex).NroMacotas - 1
For i = 1 To MAXMASCOTAS
  If UserList(userindex).MascotasIndex(i) = NpcIndex Then
     UserList(userindex).MascotasIndex(i) = 0
     UserList(userindex).MascotasType(i) = 0
     Exit For
  End If
Next i

End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer, ByVal Mascota As Integer)
    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal userindex As Integer)
On Error GoTo errhandler

   Dim MiNPC As npc
   Dim MiNPC2 As npc
   Dim MiNPC3 As npc
   Dim minpc4 As npc
   Dim asd As Integer
   MiNPC = Npclist(NpcIndex)
   MiNPC2 = Npclist(NpcIndex)
   MiNPC3 = Npclist(NpcIndex)
   
   If userindex <> 0 And Npclist(NpcIndex).NPCtype = eNPCType.ReyCastillo Then
        Call MuereRey(userindex, NpcIndex)
        Exit Sub
    End If
    
  If MiNPC3.Numero = 92 Then
        asd = Npclist(NpcIndex).MaestroUser
        UserList(asd).flags.EleDeAgua = 0
        Call QuitarNPC(NpcIndex)
        Exit Sub
    End If
    
    If MiNPC3.Numero = 93 Then
        asd = Npclist(NpcIndex).MaestroUser
        UserList(asd).flags.EleDeFuego = 0
        Call QuitarNPC(NpcIndex)
        Exit Sub
    End If
    
    If MiNPC3.Numero = 94 Then
        asd = Npclist(NpcIndex).MaestroUser
        UserList(asd).flags.EleDeTierra = 0
        Call QuitarNPC(NpcIndex)
        Exit Sub
    End If
    
   '---------------------------------------------
   '/Invocacion al matar 10 dracos :$
   '---------------------------------------------
    If MiNPC3.Numero = 564 Then 'Si muere el dragon plateado..
    
    If MurioDragon < 10 Then 'Si todavia no murieron 10..
        MurioDragon = MurioDragon + 1 'Sumamos 1
    End If
        
    If MurioDragon = 10 Then 'Si murieron 10, respawneamos.
        'Declaraciones
        Dim qwe As WorldPos
        Dim xcv As Integer
        xcv = 566
        
        qwe.Map = 96
        qwe.X = 53
        qwe.Y = 23
        Call SpawnNpc(xcv, qwe, True, False)
        Call SendData(SendTarget.toall, 0, 0, "||¡10 Dragones Plateados fueron asesinados, el Dragon de las Tinieblas ha vuelto para vengarse!. " & FONTTYPE_GUILD)
        MurioDragon = 0 'Volvemos el contador a 0..
    End If
    End If
    
    If MiNPC.Numero = 566 Then
            If UserList(userindex).Stats.TransformadoVIP = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 75 Puntos de Torneo por matar a la criatura." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 75
Call EnviarPuntos(userindex)
Else
Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 50 Puntos de Torneo por matar a la criatura." & FONTTYPE_VENENO)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 50
Call EnviarPuntos(userindex)
End If
End If
   '---------------------------------------------
   '/Invocacion al matar 10 dracos :$
   '---------------------------------------------
   
    If MiNPC3.Numero = 938 Then
    If GuardiasRey < 4 Then
        GuardiasRey = GuardiasRey + 1
    End If
        
    If GuardiasRey = 4 Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||El rey ya es vulnerable a los ataques." & FONTTYPE_ORO)
        Npclist(937).Aura = 0
    End If
    End If
    
    If (esPretoriano(NpcIndex) = 4) Then
        'seteamos todos estos 'flags' acorde para que cambien solos de alcoba
        Dim i As Integer
        Dim j As Integer
        Dim NPCI As Integer
        
        For i = 8 To 90
            For j = 8 To 90
                
                NPCI = MapData(Npclist(NpcIndex).pos.Map, i, j).NpcIndex
                If NPCI > 0 Then
                    If esPretoriano(NPCI) > 0 Then
                        Npclist(NPCI).Invent.ArmourEqpSlot = IIf(Npclist(NpcIndex).pos.X > 50, 1, 5)
                    End If
                End If
            Next j
        Next i
        Call CrearClanPretoriano(MAPA_PRETORIANO, Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y)
    ElseIf esPretoriano(NpcIndex) > 0 Then
            Npclist(NpcIndex).Invent.ArmourEqpSlot = 0
    End If
   
   'Quitamos el npc
   Call QuitarNPC(NpcIndex)
   
   If MiNPC.pos.Map = mapainvo Then MapInfo(mapainvo).criatinv = 0
   
   
    
   If userindex > 0 Then ' Lo mato un usuario?
   
   '---------------------------------------------
   'Rey :$
   '---------------------------------------------
    If UserList(userindex).pos.Map = maparey Then
    If MiNPC3.Numero = 937 Then

    'Declaraciones
    Dim Rey As WorldPos
    Dim DracoPowa As Integer
    DracoPowa = 936 'Dragon
     
    'Guardamos la posicion del rey para respawnear al dragon..
    Rey.Map = MiNPC3.pos.Map
    Rey.X = MiNPC3.pos.X
    Rey.Y = MiNPC3.pos.Y
    
        Call SpawnNpc(DracoPowa, Rey, True, False)
      End If
      
      If MiNPC3.Numero = 936 Then
    Call SendData(toall, 0, 0, "||El espiritu del rey regreso al otro mundo." & FONTTYPE_INFO)
Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 25 Puntos de Torneo por matar a la criatura." & FONTTYPE_INFO)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 25
Call EnviarPuntos(userindex)
    ReyON = 0
    GuardiasRey = 0
   End If
   End If
   
   '---------------------------------------------
   '/Rey :$
   '---------------------------------------------
   
   If UserList(userindex).pos.Map = mapainvo Then
   If MiNPC3.Numero = 911 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 912 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 913 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 914 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 915 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 916 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 917 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 918 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 919 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 920 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 921 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 922 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 923 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 924 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 925 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 926 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 927 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 928 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 929 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 930 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 931 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 932 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 933 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 934 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   If MiNPC3.Numero = 935 Then
   Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(0))
   End If
   End If
    If MiNPC3.Numero = 92 Then
        UserList(userindex).flags.EleDeAgua = 0
        Exit Sub
    End If
    
    If MiNPC3.Numero = 93 Then
        UserList(userindex).flags.EleDeFuego = 0
        Exit Sub
    End If
    
    If MiNPC3.Numero = 94 Then
        UserList(userindex).flags.EleDeTierra = 0
        Exit Sub
    End If
    
            If MiNPC.Numero = 911 Then
            If UserList(userindex).Stats.TransformadoVIP = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 45 Puntos de Torneo por matar a la criatura." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 45
Call EnviarPuntos(userindex)
Else
            Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 30 Puntos de Torneo por matar a la criatura." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 30
Call EnviarPuntos(userindex)
End If
End If

        If MiNPC.Numero = 913 Then
            If UserList(userindex).Stats.TransformadoVIP = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 18 Puntos de Torneo por matar a la criatura." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 18
Call EnviarPuntos(userindex)
Else
        Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 12 Puntos de Torneo por matar a la criatura." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 12
Call EnviarPuntos(userindex)
End If
End If

        If MiNPC.Numero = 920 Then
            If UserList(userindex).Stats.TransformadoVIP = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 45 Puntos de Torneo por matar a la criatura." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 45
Call EnviarPuntos(userindex)
Else
        Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 30 Puntos de Torneo por matar a la criatura." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 30
Call EnviarPuntos(userindex)
End If
End If

        If MiNPC.Numero = 925 Then
            If UserList(userindex).Stats.TransformadoVIP = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 37 Puntos de Torneo por matar a la criatura." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 37
Call EnviarPuntos(userindex)
Else
        Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 25 Puntos de Torneo por matar a la criatura." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 25
Call EnviarPuntos(userindex)
End If
End If

        If MiNPC.Numero = 931 Then
            If UserList(userindex).Stats.TransformadoVIP = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 75 Puntos de Torneo por matar a la criatura." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 75
Call EnviarPuntos(userindex)
Else
        Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado 50 Puntos de Torneo por matar a la criatura." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 50
Call EnviarPuntos(userindex)
End If
End If

        If MiNPC.flags.Snd3 > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & MiNPC.flags.Snd3)
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        
        'El user que lo mato tiene mascotas?
        If UserList(userindex).NroMacotas > 0 Then
            Dim T As Integer
            For T = 1 To MAXMASCOTAS
                  If UserList(userindex).MascotasIndex(T) > 0 Then
                      If Npclist(UserList(userindex).MascotasIndex(T)).TargetNPC = NpcIndex Then
                              Call FollowAmo(UserList(userindex).MascotasIndex(T))
                      End If
                  End If
            Next T
        End If
        
        '[KEVIN]
        If MiNPC.flags.ExpCount > 0 Then
            If UserList(userindex).PartyIndex > 0 Then
                Call mdParty.ObtenerExito(userindex, MiNPC.flags.ExpCount, MiNPC.pos.Map, MiNPC.pos.X, MiNPC.pos.Y)
            Else
                UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + MiNPC.flags.ExpCount
                If UserList(userindex).Stats.Exp > MAXEXP Then _
                    UserList(userindex).Stats.Exp = MAXEXP
                Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado " & MiNPC.flags.ExpCount & " puntos de experiencia." & FONTTYPE_AMARILLON)
            End If
            MiNPC.flags.ExpCount = 0
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||No has ganado experiencia al matar la criatura." & FONTTYPE_AMARILLON)
        End If
        
        '[/KEVIN]
        Call SendData(SendTarget.toindex, userindex, 0, "||Has matado a la criatura!" & FONTTYPE_FIGHT)
         Dim tmpInt As Integer
        For tmpInt = 1 To MAXUSERQUESTS
            If UserList(userindex).Stats.UserQuests(tmpInt).QuestIndex Then
                'If UserList(UserIndex).Stats.UserQuests(tmpInt).NPCsKilled = QuestList(UserList(UserIndex).Stats.UserQuests(tmpInt).QuestIndex).NpcKillIndex Then
                If MiNPC2.Numero = QuestList(UserList(userindex).Stats.UserQuests(tmpInt).QuestIndex).NpcKillIndex Then
                    UserList(userindex).Stats.UserQuests(tmpInt).NPCsKilled = UserList(userindex).Stats.UserQuests(tmpInt).NPCsKilled + 1
                End If
            End If
        Next tmpInt

   
        If UserList(userindex).Stats.NPCsMuertos < 32000 Then _
            UserList(userindex).Stats.NPCsMuertos = UserList(userindex).Stats.NPCsMuertos + 1
        
        If MiNPC.Stats.Alineacion = 0 Then
            If MiNPC.Numero = Guardias Then
               UserList(userindex).Reputacion.NobleRep = 0
                UserList(userindex).Reputacion.PlebeRep = 0
                UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 500
                If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(userindex).Reputacion.AsesinoRep = MAXREP
            End If
            If MiNPC.MaestroUser = 0 Then
                UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + vlASESINO
                If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(userindex).Reputacion.AsesinoRep = MAXREP
            End If
        ElseIf MiNPC.Stats.Alineacion = 1 Then
            UserList(userindex).Reputacion.PlebeRep = UserList(userindex).Reputacion.PlebeRep + vlCAZADOR
            If UserList(userindex).Reputacion.PlebeRep > MAXREP Then _
                UserList(userindex).Reputacion.PlebeRep = MAXREP
        ElseIf MiNPC.Stats.Alineacion = 2 Then
            UserList(userindex).Reputacion.NobleRep = UserList(userindex).Reputacion.NobleRep + vlASESINO / 2
            If UserList(userindex).Reputacion.NobleRep > MAXREP Then _
                UserList(userindex).Reputacion.NobleRep = MAXREP
        ElseIf MiNPC.Stats.Alineacion = 4 Then
            UserList(userindex).Reputacion.PlebeRep = UserList(userindex).Reputacion.PlebeRep + vlCAZADOR
            If UserList(userindex).Reputacion.PlebeRep > MAXREP Then _
                UserList(userindex).Reputacion.PlebeRep = MAXREP
        End If
        If Not Criminal(userindex) And UserList(userindex).Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(userindex)
        
        Call CheckUserLevel(userindex)
   End If ' Userindex > 0

   
   If MiNPC.MaestroUser = 0 Then
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(MiNPC)
   End If
   
   'ReSpawn o no
   Call ReSpawnNpc(MiNPC)
   
Exit Sub

errhandler:
    Call LogError("Error en MuereNpc")
    
End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
    'Clear the npc's flags
    
    With Npclist(NpcIndex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = ""
        .Attacking = 0
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .LanzaSpells = 0
        .GolpeExacto = 0
        .Invisible = 0
        .Maldicion = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
        .UseAINow = False
        .AtacaAPJ = 0
        .AtacaANPC = 0
        .AIAlineacion = e_Alineacion.ninguna
        .AIPersonalidad = e_Personalidad.ninguna
    End With
End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Contadores.Paralisis = 0
Npclist(NpcIndex).Contadores.TiempoExistencia = 0

End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Char.Body = 0
Npclist(NpcIndex).Char.CascoAnim = 0
Npclist(NpcIndex).Char.CharIndex = 0
Npclist(NpcIndex).Char.FX = 0
Npclist(NpcIndex).Char.Head = 0
Npclist(NpcIndex).Char.Heading = 0
Npclist(NpcIndex).Char.loops = 0
Npclist(NpcIndex).Char.ShieldAnim = 0
Npclist(NpcIndex).Char.WeaponAnim = 0


End Sub


Sub ResetNpcCriatures(ByVal NpcIndex As Integer)


Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroCriaturas
    Npclist(NpcIndex).Criaturas(j).NpcIndex = 0
    Npclist(NpcIndex).Criaturas(j).NpcName = ""
Next j

Npclist(NpcIndex).NroCriaturas = 0

End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)

Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroExpresiones: Npclist(NpcIndex).Expresiones(j) = "": Next j

Npclist(NpcIndex).NroExpresiones = 0

End Sub


Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)

    Npclist(NpcIndex).Attackable = 0
    Npclist(NpcIndex).CanAttack = 0
    Npclist(NpcIndex).Comercia = 0
    Npclist(NpcIndex).GiveEXP = 0
    Npclist(NpcIndex).Hostile = 0
    Npclist(NpcIndex).InvReSpawn = 0
    Npclist(NpcIndex).level = 0
    
    If Npclist(NpcIndex).MaestroUser > 0 Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)
    If Npclist(NpcIndex).MaestroNpc > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc, NpcIndex)
    
    Npclist(NpcIndex).MaestroUser = 0
    Npclist(NpcIndex).MaestroNpc = 0
    
    Npclist(NpcIndex).Mascotas = 0
    Npclist(NpcIndex).Movement = 0
    Npclist(NpcIndex).name = "NPC SIN INICIAR"
    Npclist(NpcIndex).NPCtype = 0
    Npclist(NpcIndex).Numero = 0
    Npclist(NpcIndex).Orig.Map = 0
    Npclist(NpcIndex).Orig.X = 0
    Npclist(NpcIndex).Orig.Y = 0
    Npclist(NpcIndex).PoderAtaque = 0
    Npclist(NpcIndex).PoderEvasion = 0
    Npclist(NpcIndex).Aura = 0
    Npclist(NpcIndex).pos.Map = 0
    Npclist(NpcIndex).pos.X = 0
    Npclist(NpcIndex).pos.Y = 0
    Npclist(NpcIndex).SkillDomar = 0
    Npclist(NpcIndex).Target = 0
    Npclist(NpcIndex).TargetNPC = 0
    Npclist(NpcIndex).TipoItems = 0
    Npclist(NpcIndex).Veneno = 0
    Npclist(NpcIndex).Desc = ""
    Npclist(NpcIndex).QuestNumber = 0
    Npclist(NpcIndex).TalkAfterQuest = ""
    Npclist(NpcIndex).TalkDuringQuest = ""
    
    
    Dim j As Integer
    For j = 1 To Npclist(NpcIndex).NroSpells
        Npclist(NpcIndex).Spells(j) = 0
    Next j
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)

End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer)

On Error GoTo errhandler

    Npclist(NpcIndex).flags.NPCActive = False
    
    If InMapBounds(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y) Then
        Call EraseNPCChar(SendTarget.ToMap, 0, Npclist(NpcIndex).pos.Map, NpcIndex)
    End If
    
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then
        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1
            If LastNPC < 1 Then Exit Do
        Loop
    End If
        
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1
    End If

Exit Sub

errhandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(pos As WorldPos) As Boolean
    
    If LegalPos(pos.Map, pos.X, pos.Y) Then
        TestSpawnTrigger = _
        MapData(pos.Map, pos.X, pos.Y).trigger <> 3 And _
        MapData(pos.Map, pos.X, pos.Y).trigger <> 2 And _
        MapData(pos.Map, pos.X, pos.Y).trigger <> 1
    End If

End Function

Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)
'Call LogTarea("Sub CrearNPC")
'Crea un NPC del tipo NRONPC

Dim pos As WorldPos
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long


Dim Map As Integer
Dim X As Integer
Dim Y As Integer

    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
    If nIndex = 0 Then Exit Sub
    
    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
        
        Map = OrigPos.Map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).pos = OrigPos
       
    Else
        
        pos.Map = mapa 'mapa
        altpos.Map = mapa
        
        Do While Not PosicionValida
            pos.X = RandomNumber(1, 100)    'Obtenemos posicion al azar en x
            pos.Y = RandomNumber(1, 100)    'Obtenemos posicion al azar en y
            
            Call ClosestLegalPos(pos, newpos)  'Nos devuelve la posicion valida mas cercana
            If newpos.X <> 0 Then altpos.X = newpos.X
            If newpos.Y <> 0 Then altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)
            
            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, Npclist(nIndex).flags.AguaValida) And _
               Not HayPCarea(newpos) And TestSpawnTrigger(newpos) Then
                'Asignamos las nuevas coordenas solo si son validas
                Npclist(nIndex).pos.Map = newpos.Map
                Npclist(nIndex).pos.X = newpos.X
                Npclist(nIndex).pos.Y = newpos.Y
                PosicionValida = True
            Else
                newpos.X = 0
                newpos.Y = 0
            
            End If
                
            'for debug
            Iteraciones = Iteraciones + 1
            If Iteraciones > MAXSPAWNATTEMPS Then
                If altpos.X <> 0 And altpos.Y <> 0 Then
                    Map = altpos.Map
                    X = altpos.X
                    Y = altpos.Y
                    Npclist(nIndex).pos.Map = Map
                    Npclist(nIndex).pos.X = X
                    Npclist(nIndex).pos.Y = Y
                    Call MakeNPCChar(SendTarget.ToMap, 0, Map, nIndex, Map, X, Y)
                    Exit Sub
                Else
                    altpos.X = 50
                    altpos.Y = 50
                    Call ClosestLegalPos(altpos, newpos)
                    If newpos.X <> 0 And newpos.Y <> 0 Then
                        Npclist(nIndex).pos.Map = newpos.Map
                        Npclist(nIndex).pos.X = newpos.X
                        Npclist(nIndex).pos.Y = newpos.Y
                        Call MakeNPCChar(SendTarget.ToMap, 0, newpos.Map, nIndex, newpos.Map, newpos.X, newpos.Y)
                        Exit Sub
                    Else
                        Call QuitarNPC(nIndex)
                        Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                        Exit Sub
                    End If
                End If
            End If
        Loop
        
        'asignamos las nuevas coordenas
        Map = newpos.Map
        X = Npclist(nIndex).pos.X
        Y = Npclist(nIndex).pos.Y
    End If
    
    'Crea el NPC
    Call MakeNPCChar(SendTarget.ToMap, 0, Map, nIndex, Map, X, Y)

End Sub

Sub MakeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
Dim CharIndex As Integer

    If Npclist(NpcIndex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).Char.CharIndex = CharIndex
        CharList(CharIndex) = NpcIndex
    End If
    
    MapData(Map, X, Y).NpcIndex = NpcIndex
    
    If sndRoute = SendTarget.ToMap Then
        Call ArgegarNpc(NpcIndex)
        Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "CC" & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading & "," & Npclist(NpcIndex).Char.CharIndex & "," & X & "," & Y)
    End If

End Sub

Sub ChangeNPCChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal NpcIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading)

If NpcIndex > 0 Then
    Npclist(NpcIndex).Char.Body = Body
    Npclist(NpcIndex).Char.Head = Head
    Npclist(NpcIndex).Char.Heading = Heading
    If sndRoute = SendTarget.ToMap Then
        Call SendToNpcArea(NpcIndex, "CP" & Npclist(NpcIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "CP" & Npclist(NpcIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading)
    End If
End If

End Sub

Sub EraseNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

If Npclist(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar <= 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y).NpcIndex = 0

'Actualizamos los cliente
If sndRoute = SendTarget.ToMap Then
    Call SendToNpcArea(NpcIndex, "BP" & Npclist(NpcIndex).Char.CharIndex)
Else
    Call SendData(sndRoute, sndIndex, sndMap, "BP" & Npclist(NpcIndex).Char.CharIndex)
End If

'Update la lista npc
Npclist(NpcIndex).Char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


End Sub

Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)

On Error GoTo errh
    Dim nPos As WorldPos
    nPos = Npclist(NpcIndex).pos
    Call HeadtoPos(nHeading, nPos)
    
    'Es mascota ????
    If Npclist(NpcIndex).MaestroUser > 0 Then
        ' es una posicion legal
        If LegalPos(Npclist(NpcIndex).pos.Map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida = 1) Then
        
            If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).pos.Map, nPos.X, nPos.Y) Then Exit Sub
            If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).pos.Map, nPos.X, nPos.Y) Then Exit Sub
            
Call SendToNpcArea(NpcIndex, "_" & encriptarMpNPC(Npclist(NpcIndex).Char.CharIndex, nPos.X, nPos.Y))
            
'#If SeguridadAlkon Then
'            Call SendToNpcArea(NpcIndex, "*" & Encriptacion.MoveNPCCrypt(NpcIndex, nPos.X, nPos.Y))
'#Else
'            Call SendToNpcArea(NpcIndex, "*" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
'#End If
            
            'Update map and user pos
            MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y).NpcIndex = 0
            Npclist(NpcIndex).pos = nPos
            Npclist(NpcIndex).Char.Heading = nHeading
            MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y).NpcIndex = NpcIndex
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        End If
Else ' No es mascota
        ' Controlamos que la posicion sea legal, los npc que
        ' no son mascotas tienen mas restricciones de movimiento.
        If LegalPosNPC(Npclist(NpcIndex).pos.Map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida) Then
            
            If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).pos.Map, nPos.X, nPos.Y) Then Exit Sub
            If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).pos.Map, nPos.X, nPos.Y) Then Exit Sub
            
                   Call SendToNpcArea(NpcIndex, "_" & encriptarMpNPC(Npclist(NpcIndex).Char.CharIndex, nPos.X, nPos.Y))
            
'#If SeguridadAlkon Then
'            Call SendToNpcArea(NpcIndex, "*" & Encriptacion.MoveNPCCrypt(NpcIndex, nPos.X, nPos.Y))
'#Else
'            Call SendToNpcArea(NpcIndex, "*" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
'#End If
            
            'Update map and user pos
            MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y).NpcIndex = 0
            Npclist(NpcIndex).pos = nPos
            Npclist(NpcIndex).Char.Heading = nHeading
            MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y).NpcIndex = NpcIndex
            
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
        Else
            If Npclist(NpcIndex).Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                Npclist(NpcIndex).PFINFO.PathLenght = 0
            End If
        
        End If
    End If

Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)


End Sub

Function NextOpenNPC() As Integer
'Call LogTarea("Sub NextOpenNPC")

On Error GoTo errhandler

Dim LoopC As Integer
  
For LoopC = 1 To MAXNPCS + 1
    If LoopC > MAXNPCS Then Exit For
    If Not Npclist(LoopC).flags.NPCActive Then Exit For
Next LoopC
  
NextOpenNPC = LoopC


Exit Function
errhandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal userindex As Integer)

Dim N As Integer
N = RandomNumber(1, 100)
If N < 30 Then
    UserList(userindex).flags.Envenenado = 1
    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡La criatura te ha envenenado!!" & FONTTYPE_FIGHT)
End If

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
'Crea un NPC del tipo Npcindex

Dim newpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean


Dim Map As Integer
Dim X As Integer
Dim Y As Integer
Dim it As Integer

nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

it = 0

If nIndex > MAXNPCS Then
    SpawnNpc = 0
    Exit Function
End If

Do While Not PosicionValida
        
        Call ClosestLegalPos(pos, newpos)  'Nos devuelve la posicion valida mas cercana
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If Npclist(nIndex).flags.TierraInvalida Then
            If LegalPos(newpos.Map, newpos.X, newpos.Y, True) Then _
                PosicionValida = True
        Else
            If LegalPos(newpos.Map, newpos.X, newpos.Y, False) Or LegalPos(newpos.Map, newpos.X, newpos.Y, Npclist(nIndex).flags.AguaValida) Then _
                PosicionValida = True
        End If
        
        If PosicionValida Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(nIndex).pos.Map = newpos.Map
            Npclist(nIndex).pos.X = newpos.X
            Npclist(nIndex).pos.Y = newpos.Y
        Else
            newpos.X = 0
            newpos.Y = 0
        End If
        
        it = it + 1
        
        If it > MAXSPAWNATTEMPS Then
            Call QuitarNPC(nIndex)
            SpawnNpc = 0
            Call LogError("Mas de " & MAXSPAWNATTEMPS & " iteraciones en SpawnNpc Mapa:" & pos.Map & " Index:" & NpcIndex)
            Exit Function
        End If
Loop

'asignamos las nuevas coordenas
Map = newpos.Map
X = Npclist(nIndex).pos.X
Y = Npclist(nIndex).pos.Y

'Crea el NPC
Call MakeNPCChar(SendTarget.ToMap, 0, Map, nIndex, Map, X, Y)

If FX Then
    Call SendData(SendTarget.ToNPCArea, nIndex, Map, "TW" & SND_WARP)
    Call SendData(SendTarget.ToNPCArea, nIndex, Map, "XFC" & Npclist(nIndex).Char.CharIndex & "," & FXIDs.FXWARP & "," & 0)
End If

Npclist(nIndex).Aura = 0
Call SendData(SendTarget.ToNPCArea, nIndex, Map, "AAU" & Npclist(nIndex).Char.CharIndex & "," & Npclist(nIndex).Aura)

SpawnNpc = nIndex

End Function

Sub ReSpawnNpc(MiNPC As npc)

If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.pos.Map, MiNPC.Orig)

End Sub

'Devuelve el nro de enemigos que hay en el Mapa Map
Function NPCHostiles(ByVal Map As Integer) As Integer

Dim NpcIndex As Integer
Dim cont As Integer

'Contador
cont = 0
For NpcIndex = 1 To LastNPC

    '¿esta vivo?
    If Npclist(NpcIndex).flags.NPCActive _
       And Npclist(NpcIndex).pos.Map = Map _
       And Npclist(NpcIndex).Hostile = 1 And _
       Npclist(NpcIndex).Stats.Alineacion = 2 Then
            cont = cont + 1
           
    End If
    
Next NpcIndex

NPCHostiles = cont

End Function

Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer los NPCS se deberá usar la
'nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

Dim NpcIndex As Integer
Dim npcfile As String
Dim Leer As clsIniReader

If NpcNumber > 499 Then
        'NpcFile = DatPath & "NPCs-HOSTILES.dat"
        Set Leer = LeerNPCsHostiles
Else
        'NpcFile = DatPath & "NPCs.dat"
        Set Leer = LeerNPCs
End If

NpcIndex = NextOpenNPC

If NpcIndex > MAXNPCS Then 'Limite de npcs
    OpenNPC = NpcIndex
    Exit Function
End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).name = Leer.GetValue("NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = Leer.GetValue("NPC" & NpcNumber, "Desc")

Npclist(NpcIndex).Movement = Val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

Npclist(NpcIndex).flags.AguaValida = Val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
Npclist(NpcIndex).flags.TierraInvalida = Val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
Npclist(NpcIndex).flags.Faccion = Val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))

Npclist(NpcIndex).NPCtype = Val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = Val(Leer.GetValue("NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = Val(Leer.GetValue("NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = Val(Leer.GetValue("NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = Val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = Val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = Val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile

Npclist(NpcIndex).GiveEXP = Val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP")) * 155

'Npclist(NpcIndex).flags.ExpDada = Npclist(NpcIndex).GiveEXP
Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).GiveEXP

Npclist(NpcIndex).Veneno = Val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))

Npclist(NpcIndex).flags.Domable = Val(Leer.GetValue("NPC" & NpcNumber, "Domable"))


Npclist(NpcIndex).PoderAtaque = Val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
Npclist(NpcIndex).PoderEvasion = Val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
Npclist(NpcIndex).Aura = Val(Leer.GetValue("NPC" & NpcNumber, "Aura"))

Npclist(NpcIndex).InvReSpawn = Val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
Npclist(NpcIndex).QuestNumber = Val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber"))
Npclist(NpcIndex).TalkAfterQuest = Leer.GetValue("NPC" & NpcNumber, "TalkAfterQuest")
Npclist(NpcIndex).TalkDuringQuest = Leer.GetValue("NPC" & NpcNumber, "TalkDuringQuest")
 


Npclist(NpcIndex).Stats.MaxHP = Val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = Val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = Val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = Val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.def = Val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = Val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))


Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = Val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
    Npclist(NpcIndex).Invent.Object(LoopC).ProbTirar = Val(ReadField(3, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = Val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).Amount = Val(ReadField(2, ln, 45))
Next LoopC

Npclist(NpcIndex).flags.LanzaSpells = Val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
If Npclist(NpcIndex).flags.LanzaSpells > 0 Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
    Npclist(NpcIndex).Spells(LoopC) = Val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
Next LoopC


If Npclist(NpcIndex).NPCtype = eNPCType.Entrenador Then
    Npclist(NpcIndex).NroCriaturas = Val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
    ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
    For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
        Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
        Npclist(NpcIndex).Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
    Next LoopC
End If


Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False

If Respawn Then
    Npclist(NpcIndex).flags.Respawn = Val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
Else
    Npclist(NpcIndex).flags.Respawn = 1
End If

Npclist(NpcIndex).flags.BackUp = Val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.RespawnOrigPos = Val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
Npclist(NpcIndex).flags.AfectaParalisis = Val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
Npclist(NpcIndex).flags.GolpeExacto = Val(Leer.GetValue("NPC" & NpcNumber, "GolpeExacto"))


Npclist(NpcIndex).flags.Snd1 = Val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
Npclist(NpcIndex).flags.Snd2 = Val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
Npclist(NpcIndex).flags.Snd3 = Val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

Dim aux As String
aux = Leer.GetValue("NPC" & NpcNumber, "NROEXP")
If aux = "" Then
    Npclist(NpcIndex).NroExpresiones = 0
Else
    Npclist(NpcIndex).NroExpresiones = Val(aux)
    ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
    For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
        Npclist(NpcIndex).Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
    Next LoopC
End If

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = Val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))

'Update contadores de NPCs
If NpcIndex > LastNPC Then LastNPC = NpcIndex
NumNPCs = NumNPCs + 1


'Devuelve el nuevo Indice
OpenNPC = NpcIndex

End Function


Sub EnviarListaCriaturas(ByVal userindex As Integer, ByVal NpcIndex)
  Dim SD As String
  Dim k As Integer
  SD = SD & Npclist(NpcIndex).NroCriaturas & ","
  For k = 1 To Npclist(NpcIndex).NroCriaturas
        SD = SD & Npclist(NpcIndex).Criaturas(k).NpcName & ","
  Next k
  SD = "LSTCRI" & SD
  Call SendData(SendTarget.toindex, userindex, 0, SD)
End Sub


Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)

If Npclist(NpcIndex).flags.Follow Then
  Npclist(NpcIndex).flags.AttackedBy = ""
  Npclist(NpcIndex).flags.Follow = False
  Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
  Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
Else
  Npclist(NpcIndex).flags.AttackedBy = UserName
  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = 4 'follow
  Npclist(NpcIndex).Hostile = 0
End If

End Sub

Sub FollowAmo(ByVal NpcIndex As Integer)

  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = TipoAI.SigueAmo 'follow
  Npclist(NpcIndex).Hostile = 0
  Npclist(NpcIndex).Target = 0
  Npclist(NpcIndex).TargetNPC = 0

End Sub

Public Sub MuereRey(ByVal userindex As Integer, NpcIndex As Integer)
Dim reNpcPos As WorldPos
Dim reNpcIndex As Integer
Dim Castillo As Integer
Castillo = 0
If UserList(userindex).pos.Map = MapCastilloN Then Castillo = 1
If UserList(userindex).pos.Map = MapCastilloS Then Castillo = 2
If UserList(userindex).pos.Map = MapCastilloE Then Castillo = 3
If UserList(userindex).pos.Map = MapCastilloO Then Castillo = 4
If Castillo = 0 Then Exit Sub
 
reNpcPos.Map = UserList(userindex).pos.Map
reNpcPos.X = 50
reNpcPos.Y = 34
reNpcIndex = NpcIndex
 
If Castillo = 1 Then
   CastilloNorte = Guilds(UserList(userindex).GuildIndex).GuildName
   Call SendData(toall, 0, 0, "|| El Clan " & (Guilds(UserList(userindex).GuildIndex).GuildName) & " ha conquistado el castillo Norte" & FONTTYPE_GUILD)
   Call WriteVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloNorte", Guilds(UserList(userindex).GuildIndex).GuildName)
   Call SendData(toall, 0, 0, "TW" & SND_CREACIONCLAN)
ElseIf Castillo = 2 Then
   CastilloSur = Guilds(UserList(userindex).GuildIndex).GuildName
   Call SendData(toall, 0, 0, "|| El Clan " & (Guilds(UserList(userindex).GuildIndex).GuildName) & " ha conquistado el castillo Sur" & FONTTYPE_GUILD)
   Call WriteVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloSur", Guilds(UserList(userindex).GuildIndex).GuildName)
   Call SendData(toall, 0, 0, "TW" & SND_CREACIONCLAN)
ElseIf Castillo = 3 Then
   CastilloEste = Guilds(UserList(userindex).GuildIndex).GuildName
   Call SendData(toall, 0, 0, "|| El Clan " & (Guilds(UserList(userindex).GuildIndex).GuildName) & " ha conquistado el castillo Este" & FONTTYPE_GUILD)
   Call WriteVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloEste", Guilds(UserList(userindex).GuildIndex).GuildName)
   Call SendData(toall, 0, 0, "TW" & SND_CREACIONCLAN)
ElseIf Castillo = 4 Then
   CastilloOeste = Guilds(UserList(userindex).GuildIndex).GuildName
   Call SendData(toall, 0, 0, "|| El Clan " & (Guilds(UserList(userindex).GuildIndex).GuildName) & " ha conquistado el castillo Oeste" & FONTTYPE_GUILD)
   Call WriteVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloOeste", Guilds(UserList(userindex).GuildIndex).GuildName)
   Call SendData(toall, 0, 0, "TW" & SND_CREACIONCLAN)
End If
Call QuitarNPC(NpcIndex)
Call SpawnNpc(ReyNpcN, reNpcPos, True, False)
Call SendData(SendTarget.toindex, userindex, 0, "||Has matado al rey" & FONTTYPE_INFO)
End Sub

Public Sub MuereElementalAgua(ByVal userindex As Integer, NpcIndex As Integer)
UserList(userindex).flags.EleDeAgua = 0
Call QuitarNPC(NpcIndex)
End Sub

    
Private Function encriptarMpNPC(ByVal CharIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As String
Dim Key As Byte
Key = (X Xor Y) + 8
encriptarMpNPC = Chr$((Int(CharIndex / 128) Xor X) + 32) & _
                Chr$((X Xor Key) + 8) & _
                    Chr$((Int(CharIndex Mod 128) Xor Y) + 32) & _
                        Chr$((Y Xor Key) + 8) & _
                            Chr$(Key Xor &HEC&)
End Function

