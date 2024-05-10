Attribute VB_Name = "Mod_TCP"


Option Explicit
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public nombreotro As String
Public LlegoFama As Boolean



Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
End Function

Sub HandleData(ByVal Rdata As String)
    On Error Resume Next
    
    Dim RetVal As Variant
    Dim x As Integer
    Dim Y As Integer
    Dim charindex As Integer
    Dim tempint As Integer
    Dim tempstr As String
    Dim slot As Integer
    Dim MapNumber As String
    Dim I As Integer, k As Integer
    Dim cad$, Index As Integer, m As Integer
    Dim T() As String
    
    Dim tstr As String
    Dim tstr2 As String
    
    
    Dim sData As String
    
     Rdata = DameDameX(Rdata)
        Rdata = Seventhqwedvggjfgnb(Rdata, "·")
        Rdata = Seventhqwedvggjfgnb(Rdata, "€")
        
    sData = UCase$(Rdata)
    
    If UCase$(Left$(Rdata, 2)) = "QL" Then
        Unload frmQuests
        frmQuests.lstQuests.Clear
       
        For I = 1 To 10
            tstr = ReadField(I, Right$(Rdata, Len(Rdata) - 2), Asc("-"))
           
            If tstr = "0" Then
                frmQuests.lstQuests.AddItem "NADA"
            Else
                frmQuests.lstQuests.AddItem tstr
            End If
        Next I
       
        frmQuests.Show , frmMain
        Exit Sub
    ElseIf UCase$(Left$(Rdata, 2)) = "QI" Then
        tstr = Right$(Rdata, Len(Rdata) - 2)
       
        frmQuests.lblNombre.Caption = ReadField(1, tstr, Asc("-"))
        frmQuests.lblDescripcion.Caption = ReadField(2, tstr, Asc("-"))
        frmQuests.lblCriaturas.Caption = ReadField(3, tstr, Asc("-"))
        Exit Sub
    End If
    
    Select Case sData
    
    Case "FEERASD"
    If frmRendimiento.Temuestroelcartelitoarre.value = 0 Then
    Else
       frmMuerte.Show
    End If
    Exit Sub
    
        Case "LOGINSF"            ' >>>>> LOGIN :: LOGGED
            logged = True
            UserCiego = False
            EngineRun = True
            UserDescansar = False
            Nombres = True
            If frmCrearPersonaje.Visible Then
                Unload frmPasswdSinPadrinos
                Unload frmCrearPersonaje
                Unload frmConnect
                frmMain.Show
            End If
            Call SetConnected
           
            bTecho = IIf(MapData(UserPos.x, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.x, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.x, UserPos.Y).Trigger = 4, True, False)
            Exit Sub
        Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
            Call Dialogos.BorrarDialogos
            Exit Sub
        Case "NAVEG"
            UserNavegando = Not UserNavegando
            Exit Sub
        Case "EQUIT"
            UserMontando = Not UserMontando
            Exit Sub
        Case "MEJUI" ' Graceful exit ;))
        
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            frmMain.Visible = False
            logged = False
            UserParalizado = False
            pausa = False
            UserMeditar = False
            UserDescansar = False
            UserNavegando = False
            UserMontando = False
            frmConnect.Visible = True
            Call Audio.StopWave
            bRain = False
            bFogata = False
            SkillPoints = 0
            frmMain.Label1.Visible = False
            Call Dialogos.BorrarDialogos
            For I = 1 To LastChar
                charlist(I).invisible = False
            Next I
            
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            
            bK = 0
            If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
frmMain.Socket1.Cleanup
frmConnect.MousePointer = 1
            Exit Sub
        Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = False
            Exit Sub
        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
        Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            Comerciando = False
            Exit Sub
        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
        Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
            I = 1
            Do While I <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(I) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(I)
                Else
                        frmComerciar.List1(1).AddItem "Nada"
                End If
                I = I + 1
            Loop
            Comerciando = True
            frmComerciar.Show , frmMain
            Exit Sub
        '[KEVIN]-----------------------------------------------
        '**************************************************************
               Case "INITBANKO"
            frmBanco.Show , frmMain
            Exit Sub

        Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
            Dim II As Integer
            II = 1
            Do While II <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(II) <> 0 Then
                        frmBancoObj.List1(1).AddItem Inventario.ItemName(II)
                Else
                        frmBancoObj.List1(1).AddItem "Nada"
                End If
                II = II + 1
            Loop
            
            
            I = 1
            Do While I <= UBound(UserBancoInventory)
                If UserBancoInventory(I).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(I).Name
                Else
                        frmBancoObj.List1(0).AddItem "Nada"
                End If
                I = I + 1
            Loop
            Comerciando = True
            frmBancoObj.Show , frmMain
            Exit Sub
        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
        Case "INITCOMUSU"
            If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
            If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
            
            For I = 1 To MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(I) <> 0 Then
                        frmComerciarUsu.List1.AddItem Inventario.ItemName(I)
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = Inventario.Amount(I)
                Else
                        frmComerciarUsu.List1.AddItem "Nada"
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
                End If
            Next I
            Comerciando = True
            frmComerciarUsu.Show , frmMain
        Case "FINCOMUSUOK"
            frmComerciarUsu.List1.Clear
            frmComerciarUsu.List2.Clear
            
            Unload frmComerciarUsu
            Comerciando = False
            '[/Alejo]
        Case "SFH"
            frmHerrero.Show , frmMain
            Exit Sub
        Case "SFC"
            frmCarp.Show , frmMain
            Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "SEGON" '  <--- Activa el seguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, False)
            Exit Sub
        Case "SEGOFF" ' <--- Desactiva el seguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "SEGCON" '  <--- Activa el seguro clan
            IsSeguroC = True
            Exit Sub
        Case "SEGCOFF" ' <--- Desactiva el seguro clan
            IsSeguroC = False
            Exit Sub
        Case "SEGCVCON"
            SeguroCvc = True
            Exit Sub
        Case "SEGCVCOFF"
            SeguroCvc = False
            Exit Sub
            Case "SEGONR" ' <--- Activa el seguro de resi
Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, False)
IsSeguroR = True
Exit Sub
Case "SEGOFR" ' <--- Desactiva el seguro de resu
Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, False)
IsSeguroR = False
Exit Sub
        Case "PN"     ' <--- Pierde Nobleza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, False)
            Exit Sub
    End Select

    Select Case Left(sData, 1)
        Case "+"              ' >>>>> Mover Char >>> +
            Rdata = Right$(Rdata, Len(Rdata) - 1)

            charindex = Val(ReadField(1, Rdata, Asc(",")))
            x = Val(ReadField(2, Rdata, Asc(",")))
            Y = Val(ReadField(3, Rdata, Asc(",")))


            If charlist(charindex).Fx >= 40 And charlist(charindex).Fx <= 49 Then   'si esta meditando
                charlist(charindex).Fx = 0
                charlist(charindex).FxLoopTimes = 0
            End If
            
           
            If charlist(charindex).priv = 0 Then
                Call DoPasosFx(charindex)
            End If

            Call MoveCharbyPos(charindex, x, Y)
            
            Call RefreshAllChars
            Exit Sub
        
        Case "*", "_"             ' >>>>> Mover NPC >>> *
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            
            Call LPMDESNPC(Rdata, charindex, x, Y)
            Call MoveCharbyPos(charindex, x, Y)

            Exit Sub
    
    
    End Select

    Select Case Left$(sData, 2)
        Case "AS"
            tstr = mid$(sData, 3, 1)
            k = Val(Right$(sData, Len(sData) - 3))
            
            Select Case tstr
                Case "M": UserMinMAN = Val(Right$(sData, Len(sData) - 3))
                Case "H": UserMinHP = Val(Right$(sData, Len(sData) - 3))
                Case "S": UserMinSTA = Val(Right$(sData, Len(sData) - 3))
                Case "G": UserGLD = Val(Right$(sData, Len(sData) - 3))
                Case "E": UserExp = Val(Right$(sData, Len(sData) - 3))
            End Select
            
            frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
            frmMain.ExpShp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 134)
            frmMain.lblPorcLvl.Caption = "[" & CLng(UserExp * 100 / UserPasarNivel) & "%]"
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 93)
            
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 93)
            Else
                frmMain.MANShp.Width = 0
            End If
            
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 93)
        
            frmMain.HpBar.Caption = UserMinHP & "/" & UserMaxHP
            frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            frmMain.StaBar.Caption = UserMinSTA & "/" & UserMaxSTA
            'frmMain.GldLbl.Caption = PonerPuntos(UserGLD)
            frmMain.LvlLbl.Caption = UserLvl
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
            
            If UserLvl = 50 Then
            frmMain.ExpShp.Visible = False
            Else
            frmMain.ExpShp.Visible = True
            End If
            
            
            Exit Sub
        Case "CM"              ' >>>>> Cargar Mapa :: CM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMap = ReadField(1, Rdata, 44)
            'Obtiene la version del mapa

#If SeguridadAlkon Then
            Call InitMI
#End If
            
            If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
                Open DirMapas & "Mapa" & UserMap & ".map" For Binary As #1
                Seek #1, 1
                Get #1, , tempint
                Close #1
'                If tempint = Val(ReadField(2, Rdata, 44)) Then
                    'Si es la vers correcta cambiamos el mapa
                    Call SwitchMap(UserMap)
'                Else
'                    'vers incorrecta
'                    MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
'                    Call Deinittileengine
'                    Call UnloadAllForms
'                    End
'                End If
            Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                Call DeinitTileEngine
                Call UnloadAllForms
                Call EscribirGameIni(Config_Inicio)
                End
            End If
            Exit Sub
        
        Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(UserPos.x, UserPos.Y).charindex = 0
            UserPos.x = CInt(ReadField(1, Rdata, 44))
            UserPos.Y = CInt(ReadField(2, Rdata, 44))
            MapData(UserPos.x, UserPos.Y).charindex = UserCharIndex
            charlist(UserCharIndex).Pos = UserPos
            Call MovemosUserMap
            frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.Y & ")"
            Exit Sub
        
        Case "N2" ' <<--- Npc nos impacto (Ahorramos ancho de banda)
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            I = Val(ReadField(1, Rdata, 44))
            Select Case I
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & Rdata & MENSAJE_2, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & Rdata & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            I = Val(ReadField(1, Rdata, 44))
            Select Case I
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            I = Val(ReadField(1, Rdata, 44))
            Select Case I
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "||"                 ' >>>>> Dialogo de Usuarios y NPCs :: ||
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim iuser As Integer
            iuser = Val(ReadField(3, Rdata, 176))
            
            If iuser > 0 Then
                Dialogos.CrearDialogo ReadField(2, Rdata, 176), iuser, Val(ReadField(1, Rdata, 176))
            Else
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                End If
            End If

            Exit Sub
        Case "|+"                 ' >>>>> Consola de clan y NPCs :: |+
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            iuser = Val(ReadField(3, Rdata, 176))

            If iuser = 0 Then
                If PuedoQuitarFoco And Not DialogosClanes.Activo Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                ElseIf DialogosClanes.Activo Then
                    DialogosClanes.PushBackText ReadField(1, Rdata, 126)
                End If
            End If

            Exit Sub

        Case "!!"                ' >>>>> Msgbox :: !!
            If PuedoQuitarFoco Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                frmMensaje.msg.Caption = Rdata
                frmMensaje.Show
            End If
            Exit Sub
        Case "ON"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UsersOns = Rdata
            Exit Sub
        Case "TD"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            TDAlpha = Rdata
            Exit Sub
        Case "TF"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            TDBeta = Rdata
            Exit Sub
        Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserIndex = Val(Rdata)
            Exit Sub
            Case "AF"
            If Val(Right$(Rdata, 1)) = 1 Then
                charlist(UserCharIndex).Criminal = True
            ElseIf Val(Right$(Rdata, 1)) = 0 Then
                charlist(UserCharIndex).Criminal = False
            End If
            Exit Sub

        Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserCharIndex = Val(Rdata)
            UserPos = charlist(UserCharIndex).Pos
            Call MovemosUserMap
            frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.Y & ")"
            Exit Sub
'Mithrandir
Case "CC" ' >>>>> Crear un Personaje :: CC
Rdata = Right$(Rdata, Len(Rdata) - 2)
charindex = ReadField(4, Rdata, 44)
x = ReadField(5, Rdata, 44)
Y = ReadField(6, Rdata, 44)
 
charlist(charindex).Fx = Val(ReadField(9, Rdata, 44))
charlist(charindex).FxLoopTimes = Val(ReadField(10, Rdata, 44))
charlist(charindex).Nombre = ReadField(12, Rdata, 44)
charlist(charindex).EsStatus = Val(ReadField(13, Rdata, 44))
charlist(charindex).priv = Val(ReadField(14, Rdata, 44))
'Guardamos
Call MakeChar(charindex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), x, Y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)))
Call RefreshAllChars
Exit Sub
'Mithrandir
            
            Case "AP" '[ANIM ATAK]
           Rdata = Right$(Rdata, Len(Rdata) - 2)
           
            x = Val(Rdata)
                If charlist(x).Arma.WeaponWalk(charlist(x).Heading).GrhIndex > 0 Then
                    charlist(x).Arma.WeaponWalk(charlist(x).Heading).Started = 1
                    charlist(x).Arma.WeaponAttack = GrhData(charlist(x).Arma.WeaponWalk(charlist(x).Heading).GrhIndex).NumFrames + 1
                End If
                
                Case "EM" '[ANIM ATAK ESCUDO]
            Rdata = Right$(Rdata, Len(Rdata) - 2)
           
            x = Val(Rdata)
                If charlist(x).Escudo.ShieldWalk(charlist(x).Heading).GrhIndex > 0 Then
                    charlist(x).Escudo.ShieldWalk(charlist(x).Heading).Started = 1
                    charlist(x).Escudo.ShieldAttack = GrhData(charlist(x).Escudo.ShieldWalk(charlist(x).Heading).GrhIndex).NumFrames + 1
                End If
                
                Case "NX"
Rdata = Right$(Rdata, Len(Rdata) - 2)
charindex = ReadField(1, Rdata, 44)
charlist(charindex).EsStatus = Val(ReadField(2, Rdata, 44))
charlist(charindex).Nombre = ReadField(3, Rdata, 44)
Exit Sub
            
        Case "BP"             ' >>>>> Borrar un Personaje :: BP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call EraseChar(Val(Rdata))
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Call RefreshAllChars
            Exit Sub
        Case "MP"             ' >>>>> Mover un Personaje :: MP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = Val(ReadField(1, Rdata, 44))
            
            If charlist(charindex).Fx >= 40 And charlist(charindex).Fx <= 49 Then   'si esta meditando
                charlist(charindex).Fx = 0
                charlist(charindex).FxLoopTimes = 0
            End If
            
            If charlist(charindex).priv = 0 Then
                Call DoPasosFx(charindex)
            End If
            
            Call MoveCharbyPos(charindex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44))
            
            Call RefreshAllChars
            Exit Sub
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).muerto = Val(ReadField(3, Rdata, 44)) = 500
            charlist(charindex).Body = BodyData(Val(ReadField(2, Rdata, 44)))
            charlist(charindex).Head = HeadData(Val(ReadField(3, Rdata, 44)))
            charlist(charindex).Heading = Val(ReadField(4, Rdata, 44))
            charlist(charindex).Fx = Val(ReadField(7, Rdata, 44))
            charlist(charindex).FxLoopTimes = Val(ReadField(8, Rdata, 44))
            tempint = Val(ReadField(5, Rdata, 44))
            If tempint <> 0 Then charlist(charindex).Arma = WeaponAnimData(tempint)
            tempint = Val(ReadField(6, Rdata, 44))
            If tempint <> 0 Then charlist(charindex).Escudo = ShieldAnimData(tempint)
            tempint = Val(ReadField(9, Rdata, 44))
            If tempint <> 0 Then charlist(charindex).Casco = CascoAnimData(tempint)

            Call RefreshAllChars
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            x = Val(ReadField(2, Rdata, 44))
            Y = Val(ReadField(3, Rdata, 44))
            'ID DEL OBJ EN EL CLIENTE
            MapData(x, Y).ObjGrh.GrhIndex = Val(ReadField(1, Rdata, 44))
            MapData(x, Y).ObjName = ReadField(4, Rdata, 44)
            InitGrh MapData(x, Y).ObjGrh, MapData(x, Y).ObjGrh.GrhIndex
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            x = Val(ReadField(1, Rdata, 44))
            Y = Val(ReadField(2, Rdata, 44))
            MapData(x, Y).ObjGrh.GrhIndex = 0
            MapData(x, Y).ObjName = ""
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            Dim b As Byte
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44))).Blocked = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "N~"           ' >>>>> Nombre del Mapa
         Rdata = Right$(Rdata, Len(Rdata) - 2)
            frmMain.lblMapaName.Caption = Rdata
            Exit Sub
        Case "TM"           ' >>>>> Play un MIDI :: TM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            currentMidi = Val(ReadField(1, Rdata, 45))
            
            If Musica Then
                If currentMidi <> 0 Then
                    Rdata = Right$(Rdata, Len(Rdata) - Len(ReadField(1, Rdata, 45)))
                    If Len(Rdata) > 0 Then
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Val(Right$(Rdata, Len(Rdata) - 1)))
                    Else
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
                    End If
                End If
            End If
            Exit Sub
        Case "TW"          ' >>>>> Play un WAV :: TW
            If Sound Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                 Call Audio.PlayWave(Rdata & ".wav")
            End If
            Exit Sub
            Case "PZ"          ' >>>>> Label de fuerza
Rdata = Right$(Rdata, Len(Rdata) - 2)
frmMain.Fuerza.Caption = Rdata
Exit Sub
 
Case "PX"          ' >>>>> Label de Agilidad
Rdata = Right$(Rdata, Len(Rdata) - 2)
frmMain.Agilidad.Caption = Rdata
Exit Sub

        Case "GL" 'Lista de guilds
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call frmGuildAdm.ParseGuildList(Rdata)
            Exit Sub
        Case "FO"          ' >>>>> Play un WAV :: TW
            bFogata = True
            If FogataBufferIndex = 0 Then
                FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
            End If
            Exit Sub
        Case "CA"
            CambioDeArea Asc(mid$(sData, 3, 1)), Asc(mid$(sData, 4, 1))
            Exit Sub
    End Select

    Select Case Left$(sData, 3)
    
Case "AGF"
    If CrearonTD = True Then
    CrearonTD = False
    Else
    CrearonTD = True
    End If
Exit Sub
    
Case "MMV"
If frmMain.Minimap.Visible = True Then
frmMain.Minimap.Visible = False
Else
frmMain.Minimap.Visible = True
End If
Exit Sub
    
Case "KKW"
    FeerRLZ = True
    Exit Sub
   
Case "KKQ"
    FeerRLZ = False
    Exit Sub
    
    Case "TLX"
Rdata = Right$(Rdata, Len(Rdata) - 3)
charindex = Val(ReadField(1, Rdata, 44))
charlist(charindex).Aura_Index = Val(ReadField(2, Rdata, 44))
Call InitGrh(charlist(charindex).Aura, Val(ReadField(2, Rdata, 44)))
    
Case "AAU" 'Activar aura
    AuraActivada = True
    Exit Sub
   
Case "DAU" 'Desactivar aura
    AuraActivada = False
    Exit Sub
   
Case "AUR" 'Aura
    Rdata = Right$(Rdata, Len(Rdata) - 3)
    charindex = Val(ReadField(1, Rdata, 44))
    charlist(charindex).Aura_Index = Val(ReadField(2, Rdata, 44))
    Call InitGrh(charlist(charindex).Aura, Val(ReadField(2, Rdata, 44)))
    Exit Sub
        
        Case "SPG"                  ' >>>>> Validar Cliente :: VAL
            Dim ValString As String
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            bK = CLng(ReadField(1, Rdata, Asc(",")))
            bRK = ReadField(2, Rdata, Asc(","))
            ValString = ReadField(3, Rdata, Asc(","))
            CargarCabezas
            
          
            If EstadoLogin = Normal Or EstadoLogin = CrearNuevoPj Then
                Call login
            ElseIf EstadoLogin = Dados Then
                frmCrearPersonaje.Show vbModal
            End If
            Exit Sub
            
        Case "BKW"                  ' >>>>> Pausa :: BKW
            pausa = Not pausa
            Exit Sub
 
        Case "ULZ"                  ' >>>>> Quitar Dialogo :: QDL
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Exit Sub
        Case "XFC"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).Fx = Val(ReadField(2, Rdata, 44))
            charlist(charindex).FxLoopTimes = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola GM :: AYM
            Dim n As String, n2 As String
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            n = ReadField(2, Rdata, 176)
            n2 = ReadField(1, Rdata, 176)
            frmMSG.CrearGMmSg n, n2
            frmMSG.Show , frmMain
            Exit Sub
       Case "ARM" ' fuerza y armaduras/escus/cascos en labels
        Rdata = Right$(Rdata, Len(Rdata) - 3)
       
        With frmMain
                .Arma = ReadField(1, Rdata, Asc(","))
                .Armadura = ReadField(2, Rdata, Asc(","))
                .Casco = ReadField(3, Rdata, Asc(","))
                .Escudo = ReadField(4, Rdata, Asc(","))
                .RM = ReadField(5, Rdata, Asc(","))
        End With
        
        Case "PNT"                 'Actualiza PuntosTorneo By ZaikO
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        UserPuntosTorneo = Val(ReadField(1, Rdata, 44))
        
        frmCanjes.lblptos.Caption = PonerPuntos(UserPuntosTorneo)
        Exit Sub
        
        Case "DNC"                 'Actualiza PuntosDonacion Feer~
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        UserPuntosDonacion = Val(ReadField(1, Rdata, 44))
        
        frmCanjesDonacion.lblptos.Caption = PonerPuntos(UserPuntosDonacion)
        frmCanjes.lblptosdonacion.Caption = PonerPuntos(UserPuntosDonacion)
        Exit Sub
        
        Case "VIP" 'Le cambiamos el nombre a "Desactivar VIP"
        If frmMain.hlst.List(frmMain.hlst.listIndex) = "Activar VIP" Then
        frmMain.hlst.List(frmMain.hlst.listIndex) = "Desactivar VIP"
        ElseIf frmMain.hlst.List(frmMain.hlst.listIndex) = "Desactivar VIP" Then
        frmMain.hlst.List(frmMain.hlst.listIndex) = "Activar VIP"
        End If
        Exit Sub
        
        Case "EZT"
        Rdata = Right$(Rdata, Len(Rdata) - 4)
         UserStatus = Val(ReadField(1, Rdata, 44))
        Exit Sub
        
        Case "WBP"                  ' >>>>> Actualiza Estadisticas de Usuario :: EST
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserMaxHP = Val(ReadField(1, Rdata, 44))
            UserMinHP = Val(ReadField(2, Rdata, 44))
            UserMaxMAN = Val(ReadField(3, Rdata, 44))
            UserMinMAN = Val(ReadField(4, Rdata, 44))
            UserMaxSTA = Val(ReadField(5, Rdata, 44))
            UserMinSTA = Val(ReadField(6, Rdata, 44))
            UserLvl = Val(ReadField(7, Rdata, 44))
            UserPasarNivel = Val(ReadField(8, Rdata, 44))
            UserExp = Val(ReadField(9, Rdata, 44))
                       'Standelf
            UserBOVItem = Val(ReadField(10, Rdata, 44))
            UserPuntosTorneo = Val(ReadField(11, Rdata, 44))
            UserPuntosDonacion = Val(ReadField(12, Rdata, 44))
            
            If frmBanco.Visible Then
                frmBanco.lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tienes " & UserGLD & " monedas de oro en tu billetera y en tu cuenta tienes " & UserGLDBOV & " Monedas de oro. y " & UserBOVItem & " items en tu Boveda. ¿Cómo te puedo ayudar?"
            End If
            frmMain.HpBar.Caption = UserMinHP & "/" & UserMaxHP
            frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            frmMain.StaBar.Caption = UserMinSTA & "/" & UserMaxSTA
            frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
            If UserLvl = 50 Then
            frmMain.Label2.Visible = True
            frmMain.exp.Visible = False
            frmMain.lblPorcLvl.Visible = False
            frmMain.ExpShp.Visible = False
            Else
            frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
            frmMain.exp.Visible = True
            frmMain.lblPorcLvl.Visible = True
            frmMain.ExpShp.Visible = True
            frmMain.Label2.Visible = False
            End If
            
            frmMain.ExpShp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 134)
            frmMain.lblPorcLvl.Caption = "[" & CLng(UserExp * 100 / UserPasarNivel) & "%]"
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 93)
            
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 93)
            Else
                frmMain.MANShp.Width = 0
            End If
            
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 93)
        
            'frmMain.GldLbl.Caption = PonerPuntos(UserGLD)
            frmMain.LvlLbl.Caption = UserLvl
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
        
            Exit Sub
        Case "T01"                  ' >>>>> TRABAJANDO :: TRA
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UsingSkill = Val(Rdata)
            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            End Select
            Exit Sub
        Case "FBI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            Call Inventario.SetItem(slot, ReadField(2, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), Val(ReadField(6, Rdata, 44)), Val(ReadField(7, Rdata, 44)), _
                                    Val(ReadField(8, Rdata, 44)), Val(ReadField(9, Rdata, 44)), Val(ReadField(10, Rdata, 44)), Val(ReadField(11, Rdata, 44)), ReadField(3, Rdata, 44))
            
            Exit Sub
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserBancoInventory(slot).OBJIndex = ReadField(2, Rdata, 44)
            UserBancoInventory(slot).Name = ReadField(3, Rdata, 44)
            UserBancoInventory(slot).Amount = ReadField(4, Rdata, 44)
            UserBancoInventory(slot).GrhIndex = Val(ReadField(5, Rdata, 44))
            UserBancoInventory(slot).OBJType = Val(ReadField(6, Rdata, 44))
            UserBancoInventory(slot).MaxHit = Val(ReadField(7, Rdata, 44))
            UserBancoInventory(slot).MinHit = Val(ReadField(8, Rdata, 44))
            UserBancoInventory(slot).Def = Val(ReadField(9, Rdata, 44))
        
            tempstr = ""
            
            If UserBancoInventory(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventory(slot).Amount & ") " & UserBancoInventory(slot).Name
            Else
                tempstr = tempstr & UserBancoInventory(slot).Name
            End If
            
            Exit Sub
        '************************************************************************
        '[/KEVIN]-------
        Case "ATK"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserHechizos(slot) = ReadField(2, Rdata, 44)
            If slot > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem ReadField(3, Rdata, 44)
            Else
                frmMain.hlst.List(slot - 1) = ReadField(3, Rdata, 44)
            End If
            Exit Sub
        Case "KAJ"               ' >>>>> Recibir Atributos del Personaje :: ATR
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            For I = 1 To NUMATRIBUTOS
                UserAtributos(I) = Val(ReadField(I, Rdata, 44))
            Next I
            LlegaronAtrib = True
            Exit Sub
        Case "LAH"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadField(I, Rdata, 44)
                ArmasHerrero(m) = Val(ReadField(I + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadField(I, Rdata, 44)
                ArmadurasHerrero(m) = Val(ReadField(I + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
         Case "OBR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadField(I, Rdata, 44)
                ObjCarpintero(m) = Val(ReadField(I + 1, Rdata, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
        Case "VGH"               ' >>>>> Descansar OK :: DOK
            UserDescansar = Not UserDescansar
            Exit Sub
        Case "SPL"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            For I = 1 To Val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(I + 1, Rdata, 44)
            Next I
            frmSpawnList.Show , frmMain
            Exit Sub
        Case "ERR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmConnect.MousePointer = 1
            frmPasswdSinPadrinos.MousePointer = 1
            If Not frmCrearPersonaje.Visible Then
#If UsarWrench = 1 Then
                frmMain.Socket1.Disconnect
#Else
                If frmMain.Winsock1.State <> sckClosed Then _
                    frmMain.Winsock1.Close
#End If
            End If
            MsgBox Rdata
            Exit Sub
    End Select
    
    
    Select Case Left$(sData, 4)
    Case "JKAS"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            If EHWACHIN = False Then
            EHWACHIN = True
            End If
            Call EHWACHO
            Exit Sub
    Case "PCCC"
            Dim Caption As String
            Dim Nomvre As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Caption = ReadField(1, Rdata, 44)
            Nomvre = ReadField(2, Rdata, 44)
            Call frmCaptions.Show
            frmCaptions.List1.AddItem Caption
            frmCaptions.Caption = "Captions de " & Nomvre
        Case "PCCP"
            frmCaptions.List1.Clear
            frmCaptions.Caption = ""
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            charindex = Val(ReadField(1, Rdata, 44))
            Call frmCaptions.Listar(charindex)
            Exit Sub
            Case "MENU"
                Dim esgm As Byte
                Rdata = Right$(Rdata, Len(Rdata) - 4)
                nombreotro = ReadField(1, Rdata, 44)
                esgm = ReadField(2, Rdata, 44)
                If esgm > 0 Then
                frmMenuGM.Show , frmMain
                Else
                frmMen.Show , frmMain
                End If
                Exit Sub
        Case "PART"
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ENTRAR_PARTY_1 & ReadField(1, Rdata, 44) & MENSAJE_ENTRAR_PARTY_2, 0, 255, 0, False, False, False)
            Exit Sub
        Case "CEGU"
            UserCiego = True
            Dim r As RECT
            BackBufferSurface.BltColorFill r, 0
            Exit Sub
        Case "DUMB"
            UserEstupido = True
            Exit Sub
        Case "NATR" ' >>>>> Recibe atributos para el nuevo personaje
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserAtributos(1) = ReadField(1, Rdata, 44)
            UserAtributos(2) = ReadField(2, Rdata, 44)
            UserAtributos(3) = ReadField(3, Rdata, 44)
            UserAtributos(4) = ReadField(4, Rdata, 44)
            UserAtributos(5) = ReadField(5, Rdata, 44)
            
            frmCrearPersonaje.lbFuerza.Caption = UserAtributos(1)
            frmCrearPersonaje.lbInteligencia.Caption = UserAtributos(2)
            frmCrearPersonaje.lbAgilidad.Caption = UserAtributos(3)
            frmCrearPersonaje.lbCarisma.Caption = UserAtributos(4)
            frmCrearPersonaje.lbConstitucion.Caption = UserAtributos(5)
            
            Exit Sub
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Call InitCartel(ReadField(1, Rdata, 176), CInt(ReadField(2, Rdata, 176)))
            Exit Sub
        Case "EFYK"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            Rdata = Right(Rdata, Len(Rdata) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).Name = ReadField(1, Rdata, 44)
            NPCInventory(NPCInvDim).Amount = ReadField(2, Rdata, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(3, Rdata, 44)
            NPCInventory(NPCInvDim).GrhIndex = ReadField(4, Rdata, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(5, Rdata, 44)
            NPCInventory(NPCInvDim).OBJType = ReadField(6, Rdata, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(7, Rdata, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(8, Rdata, 44)
            NPCInventory(NPCInvDim).Def = ReadField(9, Rdata, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(10, Rdata, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(11, Rdata, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(12, Rdata, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(13, Rdata, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(14, Rdata, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(15, Rdata, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(16, Rdata, 44)
            frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
            Exit Sub

        Case "YGIJ"             ' >>>>> Recibe Fama de Personaje :: FAMA
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserReputacion.AsesinoRep = Val(ReadField(1, Rdata, 44))
            UserReputacion.BandidoRep = Val(ReadField(2, Rdata, 44))
            UserReputacion.BurguesRep = Val(ReadField(3, Rdata, 44))
            UserReputacion.LadronesRep = Val(ReadField(4, Rdata, 44))
            UserReputacion.NobleRep = Val(ReadField(5, Rdata, 44))
            UserReputacion.PlebeRep = Val(ReadField(6, Rdata, 44))
            UserReputacion.Promedio = Val(ReadField(7, Rdata, 44))
            LlegoFama = True
            Exit Sub
        Case "KIDX" ' >>>>>> Mini Estadisticas :: MEST
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            With UserEstadisticas
                .CiudadanosMatados = Val(ReadField(1, Rdata, 44))
                .CriminalesMatados = Val(ReadField(2, Rdata, 44))
                .UsuariosMatados = Val(ReadField(3, Rdata, 44))
                .NpcsMatados = Val(ReadField(4, Rdata, 44))
                .Clase = ReadField(5, Rdata, 44)
                .PenaCarcel = Val(ReadField(6, Rdata, 44))
                .NeutralesMatados = Val(ReadField(7, Rdata, 44))
            End With
            Exit Sub
        Case "UIVC"             ' >>>>> Subir Nivel :: SUNI
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            SkillPoints = SkillPoints + Val(Rdata)
            frmMain.Label1.Visible = True
            Exit Sub
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            AddtoRichTextBox frmMain.RecTxt, MENSAJE_NENE & Rdata, 255, 255, 255, 0, 0
            Exit Sub
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmMSG.List1.AddItem Rdata
            Exit Sub
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            frmMSG.Show , frmMain
            Exit Sub
        Case "FMSG"             ' >>>>> Foros :: FMSG
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmForo.List.AddItem ReadField(1, Rdata, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show , frmMain
            End If
            Exit Sub
    End Select

    Select Case Left$(sData, 5)
    
        Case UCase$(Chr$(110)) & mid$("PEDOP", 4, 1) & Right$("akV", 1) & "E" & Trim$(Left$("  RS", 3))
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).invisible = (Val(ReadField(2, Rdata, 44)) = 1)
            
#If SeguridadAlkon Then
            If (10 * Val(ReadField(2, Rdata, 44)) = 10) Then
                Call MI(CualMI).SetInvisible(charindex)
            Else
                Call MI(CualMI).ResetInvisible(charindex)
            End If
#End If

            Exit Sub
        Case "DAMEX"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            With frmCrearPersonaje
                If .Visible Then
                    .lbFuerza.Caption = ReadField(1, Rdata, 44)
                    .lbAgilidad.Caption = ReadField(2, Rdata, 44)
                    .lbInteligencia.Caption = ReadField(3, Rdata, 44)
                    .lbCarisma.Caption = ReadField(4, Rdata, 44)
                    .lbConstitucion.Caption = ReadField(5, Rdata, 44)
                End If
            End With
            
            Exit Sub
        Case "PEDOP"            ' >>>>> Meditar OK :: MEDOK
        UserMeditar = Not UserMeditar
            Call Audio.PlayWave("604.wav") 'Metemos sonido
        Exit Sub
        Case "SOUND" 'Que mierda hace esto aca (?, con esto vamos a parar todos los sonidos que se les cante la chota desde el servidor.
            Audio.StopWave 'Frenamos el sonido
        Exit Sub
    End Select

    Select Case Left(sData, 6)

        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "LLSIKS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            For I = 1 To NUMSKILLS
                UserSkills(I) = Val(ReadField(I, Rdata, 44))
            Next I
            LlegaronSkills = True
            Exit Sub
        Case "LSTCRI"
            Rdata = Right(Rdata, Len(Rdata) - 6)
            For I = 1 To Val(ReadField(1, Rdata, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(I + 1, Rdata, 44)
            Next I
            frmEntrenador.Show , frmMain
            Exit Sub
    End Select
    
    Select Case Left$(sData, 7)
        Case "RESPUES"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            TieneParaResponder = True
            frmRespuestaGM.Label1.Caption = Rdata
        Case "NEWSOSM"
        Debug.Print "Me llego mensaje"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            MensajesNumber = MensajesNumber + 1
            MensajesSOS(MensajesNumber).TIPO = ReadField(1, Rdata, Asc(","))
            MensajesSOS(MensajesNumber).Autor = ReadField(2, Rdata, Asc(","))
            MensajesSOS(MensajesNumber).Contenido = ReadField(3, Rdata, Asc(","))
        Case "GUILDNE"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildNews.ParseGuildNews(Rdata)
            Exit Sub
        Case "PEACEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEPR"  'lista de prop de alianzas
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParseAllieOffers(Rdata)
        Case "PEACEPR"  'lista de prop de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
        Case "CHRINFO"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmCharInfo.parseCharInfo(Rdata)
            Exit Sub
        Case "IREDAEL"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildLeader.ParseLeaderInfo(Rdata)
            Exit Sub
        Case "LEADSUB"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildLeader.ParseSubLeaderInfo(Rdata)
            Exit Sub
        Case "CLANDET"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildBrief.ParseGuildInfo(Rdata)
            Exit Sub
        Case "CLANDETSUB"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildBrief.ParseSubGuildInfo(Rdata)
            Exit Sub
        Case "SHOWFUN"
            CreandoClan = True
            frmGuildFoundation.Show , frmMain
            Exit Sub
        Case "PARADOK"         ' >>>>> Paralizar OK :: PARADOK
            UserParalizado = Not UserParalizado
            Exit Sub
        Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                I = 1
                Do While I <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(I) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(I)
                    Else
                        frmComerciar.List1(1).AddItem "Nada"
                    End If
                    I = I + 1
                Loop
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                    frmComerciar.List1(0).listIndex = frmComerciar.LastIndex1
                Else
                    frmComerciar.List1(1).listIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************
        Case "BANCOOK"           ' Banco OK :: BANCOOK
            If frmBancoObj.Visible Then
                I = 1
                Do While I <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(I) <> 0 Then
                            frmBancoObj.List1(1).AddItem Inventario.ItemName(I)
                    Else
                            frmBancoObj.List1(1).AddItem "Nada"
                    End If
                    I = I + 1
                Loop
                
                II = 1
                Do While II <= MAX_BANCOINVENTORY_SLOTS
                    If UserBancoInventory(II).OBJIndex <> 0 Then
                            frmBancoObj.List1(0).AddItem UserBancoInventory(II).Name
                    Else
                            frmBancoObj.List1(0).AddItem "Nada"
                    End If
                    II = II + 1
                Loop
                
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                        frmBancoObj.List1(0).listIndex = frmBancoObj.LastIndex1
                Else
                        frmBancoObj.List1(1).listIndex = frmBancoObj.LastIndex2
                End If
            End If
            Exit Sub
        '[/KEVIN]************************************************************************
        '----------------------------------------------------------------------------------
               Case "TRAVELS"
            Call frmViajes.Show(vbModeless, frmMain)
            Exit Sub
        Case "CONSULT"
                frmConsultas.Show
        Exit Sub
        Case "ABPANEL"
            frmPanelGm.Show vbModal, frmMain
            Exit Sub
        Case "LISTUSU"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            T = Split(Rdata, ",")
            If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear
                For I = LBound(T) To UBound(T)
                    'frmPanelGm.cboListaUsus.AddItem IIf(Left(t(i), 1) = " ", Right(t(i), Len(t(i)) - 1), t(i))
                    frmPanelGm.cboListaUsus.AddItem T(I)
                Next I
                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.listIndex = 0
            End If
            Exit Sub
    End Select
    
    '[Alejo]
    Select Case UCase$(Left$(Rdata, 9))
        Case "COMUSUINV"
            Rdata = Right$(Rdata, Len(Rdata) - 9)
            OtroInventario(1).OBJIndex = ReadField(2, Rdata, 44)
            OtroInventario(1).Name = ReadField(3, Rdata, 44)
            OtroInventario(1).Amount = ReadField(4, Rdata, 44)
            OtroInventario(1).Equipped = ReadField(5, Rdata, 44)
            OtroInventario(1).GrhIndex = Val(ReadField(6, Rdata, 44))
            OtroInventario(1).OBJType = Val(ReadField(7, Rdata, 44))
            OtroInventario(1).MaxHit = Val(ReadField(8, Rdata, 44))
            OtroInventario(1).MinHit = Val(ReadField(9, Rdata, 44))
            OtroInventario(1).Def = Val(ReadField(10, Rdata, 44))
            OtroInventario(1).Valor = Val(ReadField(11, Rdata, 44))
            
            frmComerciarUsu.List2.Clear
            
            frmComerciarUsu.List2.AddItem OtroInventario(1).Name
            frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount
            
            frmComerciarUsu.lblEstadoResp.Visible = False
    End Select

End Sub

Sub SendData(ByVal sdData As String)

    sdData = asdasjfjlkawqwr(sdData, "´")
        sdData = asdasjfjlkawqwr(sdData, "*")
        sdData = Encriptar(sdData)
        
        
#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then Exit Sub
#Else
    If frmMain.Winsock1.State <> sckConnected Then Exit Sub
#End If

    Dim AuxCmd As String
    AuxCmd = UCase$(Left$(sdData, 5))
    
    sdData = sdData & ENDC

    If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then
        Exit Sub
    ElseIf Len(sdData) > 300 And AuxCmd <> "DEMSG" Then
        Exit Sub
    End If

#If UsarWrench = 1 Then
    Call frmMain.Socket1.Write(sdData, Len(sdData))
#Else
    Call frmMain.Winsock1.SendData(sdData)
#End If

End Sub

Sub login()

    If EstadoLogin = Normal Then
     
    SendData ("UI!GSZ" & Val(HDSerial))
        SendData ("HJ6MXN" & UserName & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision)
    ElseIf EstadoLogin = CrearNuevoPj Then
    SendData ("UI!GSZ" & Val(HDSerial))
    
       SendData "LK!JCM" & UserName & "," & UserPassword & "," & App.Major & "." & App.Minor _
        & "." & App.Revision & "," & UserRaza & "," & UserSexo & "," & UserClase & "," & UserEmail _
        & "," & "" & "," & UserPin & "," & MiCabeza
    
    
    End If
    
End Sub

Function Ofdjclsdkf(ByVal s As String) As String
Dim I As Integer, r As String
Dim C1 As Integer, C2 As Integer
Dim p As String
r = ""
p = "PzH642!5hQ19!"
If Len(p) > 0 Then
For I = 1 To Len(s)
C1 = Asc(mid(s, I, 1))
If I > Len(p) Then
C2 = Asc(mid(p, I Mod Len(p) + 1, 1))
Else
C2 = Asc(mid(p, I, 1))
End If
C1 = C1 + C2 + 64
If C1 > 255 Then C1 = C1 - 256
r = r + Chr(C1)
Next I
Else
r = s
End If
Ofdjclsdkf = r
End Function

Function Lfdjcnzsmfdd(ByVal s As String) As String
Dim I As Integer, r As String
Dim C1 As Integer, C2 As Integer
Dim p As String
p = "PzH642!5hQ19!"
r = ""
If Len(p) > 0 Then
For I = 1 To Len(s)
C1 = Asc(mid(s, I, 1))
If I > Len(p) Then
C2 = Asc(mid(p, I Mod Len(p) + 1, 1))
Else
C2 = Asc(mid(p, I, 1))
End If
C1 = C1 - C2 - 64
If Sgn(C1) = -1 Then C1 = 256 + C1
r = r + Chr(C1)
Next I
Else
r = s
End If
Lfdjcnzsmfdd = r
End Function


Public Function asdasdsad(sTexto As String) As String
Dim I As Integer
Dim CodeAscii As Integer 'Almacena el codigo Ascii de la letra
Dim sLetra As String 'Almacena una letra
    'Bucle que recorre cada letra del sTexto
    For I = 1 To Len(sTexto)
        sLetra = mid(sTexto, I, 1) 'Almacena la letra
            CodeAscii = ((Asc(sLetra) + 758) - 148) 'Obtiene el Ascii del sLetra
            If CodeAscii < 100 Then  'Si es menor que 100
                asdasdsad = asdasdsad & "0" & CodeAscii 'Imprime un 0 delante para que tenga 3 caracteres
            Else
                asdasdsad = asdasdsad & CodeAscii 'Lo deja talcual
            End If
    DoEvents 'Realiza cada evento
    Next I
End Function 'Fin de la funcion


Private Sub LPMDESNPC(ByVal inputStr As String, ByRef charindex As Integer, ByRef x As Integer, ByRef Y As Integer)
    Dim key As Byte
    key = Asc(mid$(inputStr, 5, 1)) Xor &HEC&
    x = (Asc(mid$(inputStr, 2, 1)) - 8) Xor key
    Y = (Asc(mid$(inputStr, 4, 1)) - 8) Xor key
    charindex = ((Asc(mid$(inputStr, 1, 1)) - 32) Xor x) * 128 + (Asc(mid$(inputStr, 3, 1)) - 32) Xor Y
End Sub


