Attribute VB_Name = "ModCuentas"
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'&&& Sistema de Cuentas by SheKme &&&
'&&& Geo AO 2.0      Mayo de 2010 &&&
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
 
'Declaracion de variables
Public DirAccount As String
 
Option Explicit
 
'----------------
'STRUCTURE SYSTEM
'----------------
Function Start_Account_System()
'$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$ Author: SheKme $$$$
'$$$$   14/05/2010   $$$$
'$$$$$$$$$$$$$$$$$GeoAO$$
       
'Account's folder
DirAccount = App.Path + "\Cuentas\"
       
End Function
 
Sub LoginAccount(ByVal UserIndex As Integer, NameAccount As String, PasswordAccount As String)
'$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$ Author: SheKme $$$$
'$$$$   14/05/2010   $$$$
'$$$$$$$$$$$$$$$$$GeoAO$$
 
On Error GoTo ErrAcc
 
'...Account's login system...
 
Dim PackToSend As String, _
    Slot As Byte, _
    TotalSlots As Byte, _
    AccountData As String
   
AccountData = DirAccount & NameAccount & ".acc"
 
'if not exist the account...
If Not FileExist(AccountData, vbNormal) Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa cuenta '" & NameAccount & "' no existe.")
    Exit Sub
End If
 
'If the password is incorrect...
If GetVar(AccountData, "INIT", "Password") <> PasswordAccount Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa contraseña es incorrecta.")
    Exit Sub
End If
 
'If the account has bloqued...
If GetVar(AccountData, "INIT", "Bloq") = "True" Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa cuenta esta bloqueada.")
    Exit Sub
End If
 
'If all right, start the load process!
 
TotalSlots = Val(GetVar(AccountData, "INIT", "TotalPJs"))
 
'If the account have zero chars, then...
If TotalSlots = 0 Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "LK" & 0)
    Exit Sub
End If
 
'Else...
PackToSend = TotalSlots
 
'Charging the package to will be send...
For Slot = 1 To TotalSlots
    PackToSend = PackToSend & "," & GetVar(AccountData, "PJs", "Name" & Slot)
Next Slot
 
'Send the package...
Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "LK" & PackToSend)
 
'End of Function
Exit Sub
 
'If an error have ocurred...
ErrAcc:
Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRSe ha producido un error en la carga de la cuenta. Contactese con el Soporte de Geo AO")
'End of Function
End Sub
 
Sub MakeAccount(ByVal UserIndex As Integer, NameAccount As String, PasswordAccount As String, EmailAccount As String)
'$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$ Author: SheKme $$$$
'$$$$   14/05/2010   $$$$
'$$$$$$$$$$$$$$$$$GeoAO$$
 
On Error GoTo ErrAcc
 
'...Account's creation system...
 
Dim AccountData As String, _
    CodeRec As Long
   
AccountData = DirAccount & NameAccount & ".acc"
 
'If the account exist...
If FileExist(AccountData, vbArchive) Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa cuenta '" & NameAccount & "' ya existe. Intenta con otro nombre.")
    Exit Sub
End If
 
'If all right, start the creation process!
 
CodeRec = RandomNumber(1000, 9999)
Call WriteVar(AccountData, "INIT", "Password", PasswordAccount)
Call WriteVar(AccountData, "INIT", "TotalPJs", 0)
Call WriteVar(AccountData, "INIT", "Email", EmailAccount)
Call WriteVar(AccountData, "PJs", "Name1", "Nothing")
Call WriteVar(AccountData, "PJs", "Name2", "Nothing")
Call WriteVar(AccountData, "PJs", "Name3", "Nothing")
Call WriteVar(AccountData, "PJs", "Name4", "Nothing")
Call WriteVar(AccountData, "PJs", "Name5", "Nothing")
Call WriteVar(AccountData, "PJs", "Name6", "Nothing")
Call WriteVar(AccountData, "PJs", "Name7", "Nothing")
Call WriteVar(AccountData, "PJs", "Name8", "Nothing")
Call WriteVar(AccountData, "INIT", "CodeRec", CodeRec)
Call WriteVar(AccountData, "INIT", "Bloq", "False")
 
Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "CK" & NameAccount & "," & CodeRec)
 
'End of Function
Exit Sub
 
'If an error have ocurred...
ErrAcc:
Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRSe ha producido un error en la creacion de la cuenta. Contactese con el Soporte de Geo AO")
'End of Function
End Sub
 
Sub WriteNewCharInAccount(ByVal UserIndex As Integer, NameAccount As String, NameChar As String)
'$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$ Author: SheKme $$$$
'$$$$   14/05/2010   $$$$
'$$$$$$$$$$$$$$$$$GeoAO$$
 
On Error GoTo ErrAcc
 
'...Account's new char write system...
 
Dim TotalSlots As Byte, _
    AccountData As String
   
AccountData = DirAccount & NameAccount & ".acc"
   
'Load the total chars in memory for save the new value...
TotalSlots = Val(GetVar(AccountData, "INIT", "TotalPJs"))
TotalSlots = TotalSlots + 1
Call WriteVar(AccountData, "INIT", "TotalPJs", TotalSlots)
Call WriteVar(AccountData, "PJs", "Name" & TotalSlots, NameChar)
 
'End of Function
Exit Sub
 
'If an error have ocurred...
ErrAcc:
Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRSe ha producido un error en la actualizacion de la cuenta. Contactese con el Soporte de Geo AO")
'End of Function
End Sub
 
Sub KillCharInAccount(ByVal UserIndex As Integer, NameAccount As String, NameChar As String, KillChrArchive As Boolean)
'$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$ Author: SheKme $$$$
'$$$$   14/05/2010   $$$$
'$$$$$$$$$$$$$$$$$GeoAO$$
 
On Error GoTo ErrAcc
 
'...Account's kill char system...
 
Dim TotalSlots As Byte, _
    AccountData As String, _
    Slot As Byte, _
    LoopC As Byte, _
    TempMemory As String
 
AccountData = DirAccount & NameAccount & ".acc"
 
'Load the TotalPJs's value...
TotalSlots = Val(GetVar(AccountData, "INIT", "TotalPJs"))
 
'Search the char's position in the Account
For LoopC = 1 To TotalSlots
    TempMemory = GetVar(AccountData, "PJs", "Name" & LoopC)
    If TempMemory = NameChar Then
        Slot = LoopC
        Exit For
    End If
Next LoopC
 
'Relocation of chars...
For LoopC = Slot + 1 To TotalSlots
    TempMemory = GetVar(AccountData, "PJs", "Name" & LoopC)
    Call WriteVar(AccountData, "PJs", "Name" & (LoopC - 1), TempMemory)
Next LoopC
 
Call WriteVar(AccountData, "PJs", "Name" & TotalSlots, "Nothing")
 
'Decrement...
Call WriteVar(AccountData, "INIT", "TotalPJs", TotalSlots - 1)
 
'Kill the CHR archive (bool)
If KillChrArchive Then Kill (App.Path + "\Charfile\" & NameChar & ".chr")
 
'End of Function
Exit Sub
 
'If an error have ocurred...
ErrAcc:
Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRSe ha producido un error en la actualizacion de la cuenta. Contactese con el Soporte de Geo AO")
'End of Function
End Sub
 
Sub CharMigrationAccountToAccount(ByVal UserIndex As Integer, NameAccountOrigen As String, NameAccountDestino As String, NameChar As String)
'$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$ Author: SheKme $$$$
'$$$$   14/05/2010   $$$$
'$$$$$$$$$$$$$$$$$GeoAO$$
 
'...Account's migration char system...
 
KillCharInAccount UserIndex, NameAccountOrigen, NameChar, False
WriteNewCharInAccount UserIndex, NameAccountDestino, NameChar
 
'End of Function
End Sub
 
Sub RescueAccount(ByVal UserIndex As Integer, NameAccount As String, PasswordAccount As String, EmailAccount As String, CodeNumber As Long)
'$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$ Author: SheKme $$$$
'$$$$   16/05/2010   $$$$
'$$$$$$$$$$$$$$$$$GeoAO$$
 
On Error GoTo ErrAcc
 
'...Account's rescue system...
 
Dim AccountData As String, _
    EmailAcc As String, _
    CodeNumAcc As Long
 
AccountData = DirAccount & NameAccount & ".acc"
 
'Check the email...
EmailAcc = GetVar(AccountData, "INIT", "Email")
If EmailAccount <> EmailAcc Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERREl email es incorrecto.")
    Exit Sub
End If
   
'Check the code...
CodeNumAcc = Val(GetVar(AccountData, "INIT", "NumRec"))
If CodeNumAcc <> CodeNumber Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERREl codigo de recuperacion es incorrecto.")
    Exit Sub
End If
 
'If the account has bloqued...
If GetVar(AccountData, "INIT", "Bloq") = "True" Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa cuenta se encuentra bloqueada.")
    Exit Sub
End If
 
'If all right, then...
Call WriteVar(AccountData, "INIT", "Password", PasswordAccount)
Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRSe ha recuperado la cuenta '" & NameAccount & "'. Ahora puedes ingresar a ella con la nueva contraseña.")
 
'End of Function
Exit Sub
 
'If an error have ocurred...
ErrAcc:
Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRSe ha producido un error en la recuperacion de la cuenta. Contactese con el Soporte de Geo AO")
'End of Function
End Sub
 
Sub KillAccount(ByVal UserIndex As Integer, NameAccount As String, PasswordAccount As String, EmailAccount As String, CodeNumber As Long)
'$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$ Author: SheKme $$$$
'$$$$   17/05/2010   $$$$
'$$$$$$$$$$$$$$$$$GeoAO$$
 
On Error GoTo ErrAcc
 
'...Account's killer system...
 
Dim NameKillAcc As String, _
    PasswordKillAcc As String, _
    EmailKillAcc As String, _
    CodeKillAcc As Long, _
    AccountData As String
   
AccountData = DirAccount & NameAccount & ".acc"
 
'Check the Name...
If Not FileExist(AccountData, vbNormal) Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa cuenta '" & NameAccount & "' no existe.")
    Exit Sub
End If
 
'Check the password...
PasswordKillAcc = GetVar(AccountData, "INIT", "Password")
If PasswordKillAcc <> PasswordAccount Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa contraseña de la cuenta es incorrecta.")
    Exit Sub
End If
 
'Check the Email...
EmailKillAcc = GetVar(AccountData, "INIT", "Email")
If EmailKillAcc <> EmailAccount Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERREl email es incorrecto.")
    Exit Sub
End If
 
'Check the Code of Rescue...
CodeKillAcc = Val(GetVar(AccountData, "INIT", "CodeRec"))
If CodeKillAcc <> CodeNumber Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERREl codigo de rescate es incorrecto.")
    Exit Sub
End If
 
'If all right, then kill the account!
 
Kill AccountData
 
Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa cuenta '" & NameAccount & "' ha sido borrada con exito.")
 
'End of Function
Exit Sub
 
'If an error have ocurred...
ErrAcc:
Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRSe ha producido un error en el borrado de la cuenta. Contactese con el Soporte de Geo AO")
'End of Function
End Sub
 
Sub LockAccount(ByVal UserIndex As Integer, NameChar As String, Optional Bloq As Boolean = False, Optional UnBloq As Boolean = False)
'$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$ Author: SheKme $$$$
'$$$$   19/05/2010   $$$$
'$$$$$$$$$$$$$$$$$GeoAO$$
 
On Error GoTo ErrAcc
 
Dim CharRouteData As String
    NameAccount As String
    AccountData As String
   
'Load the route of .chr file...
CharRouteData = App.Path + "\Charfile\" & NameChar & ".chr"
 
 
'If the file .chr not exist...
If Not FileExist(CharRouteData, vbNormal) Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERREl personaje es invalido.")
    Exit Sub
End If
 
'Load the account's name...
NameAccount = GetVar(CharRouteData, "ACCOUNT", "Name")
       
'Make the route of .acc file...
AccountData = DirAccount & NameAccount & ".acc"
 
'If the file .acc not exist...
If Not FileExist(CharRouteData, vbNormal) Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa cuenta establecida en la base de datos del personaje '" & NameChar & "' es erronea. Contactese con el Soporte de Geo AO")
    Exit Sub
End If
 
'If the account has bloqued...
If Bloq = True And GetVar(AccountData, "INIT", "Bloq") = "True" And UnBloq = False Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa cuenta '" & NameAccount & "' del personaje '" & NameChar & "' ya se encuentra bloqueada.")
    Exit Sub
End If
 
'If the account hasn't bloqued...
If UnBloq = True And GetVar(AccountData, "INIT", "Bloq") = "False" And Bloq = False Then
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa cuenta '" & NameAccount & "' del personaje '" & NameChar & "' ya se encuentra desbloqueada.")
    Exit Sub
End If
 
'If all right, then lock the account!
If Bloq And Not UnBloq Then
    Call WriteVar(AccountData, "INIT", "Bloq", "True")
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa cuenta '" & NameAccount & "' del personaje '" & NameChar & "' ha sido bloqueada con exito.")
    Exit Sub
End If
 
'If all right, then unlock the account!
If UnBloq And Not Bloq Then
    Call WriteVar(AccountData, "INIT", "Bloq", "False")
    Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRLa cuenta '" & NameAccount & "' del personaje '" & NameChar & "' ha sido desbloqueada con exito.")
    Exit Sub
End If
 
'End the Function
Exit Sub
 
'If an error have ocurred...
ErrAcc:
Call TCP.PACK_DATA_SEND_TO_INDEX(UserIndex, "ERRSe ha producido un error al intentar bloquear/desbloquear la cuenta. Por favor, contactese con el Soporte de Geo AO")
 
'End the Function
End Sub

