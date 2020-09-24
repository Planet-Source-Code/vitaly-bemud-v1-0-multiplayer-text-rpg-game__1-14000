Attribute VB_Name = "mdlPcCreation"
'****************************************************************************
'*  Original BeMud copyright (C)1999-2000 by Vitaly Belman                  *
'*                                                                          *
'*  In order to use any part of BeMud, you must comply with                 *
'*  both the original BeMud license in 'license.doc'.  As stated in the     *
'*  'License.doc', you may not remove any orignal copyright.                *
'*                                                                          *
'*   BeMud is copyright 1999 - Vitaly Belman                                *
'*   BeMud is:                                                              *
'*       Vitaly Belman (vitali@actcom.co.il)                                *
'*       ICQ: 1912453                                                       *
'*                                                                          *
'*   By using this code, you have agreed to follow the terms of the         *
'*   BeMud license, in the file license.doc                                 *
'****************************************************************************
Option Explicit
'Character creation
Sub DoCreation(Index As Integer, DataR$)
    Dim LowData$, Data$
    Data = DataR
    Char(Index).Data = ""
    LowData = LCase(Data)
    Select Case Char(Index).GameState
    Case "Name"
'\B/--------------------------Checking if the name is legal--------------------------
        If InStr(Data, " ") = 0 And Len(Data) > 0 And LettersOnly(Data) And _
           InStr("," & ForbiddenNames & ",", Data) = 0 Then
'/E\--------------------------Checking if the name is legal--------------------------
            Char(Index).Name = StrConv(Data, vbProperCase) 'Fixing Name's case
            If Not dbIsExist(CharRecord.SnapShot, CharRecord.Table, Index, "Name", Data) Then
                Send Index, "This name doesn't exist. Are you sure " & QteMe(Char(Index).Name, "'") & " would be a good name?"
                Char(Index).GameState = "NameConfirm"
            Else
                Send Index, "This character already exists. Enter your password, please: ", , , ""
                Char(Index).GameState = "PasswordCheck"
            End If
        Else
            Send Index, "Illegal name. Type again: ", , , ""
        End If
    Case "NameConfirm"
        Select Case LowData
        Case "y", "yes"
'< Saving to database
            Send Index, "Please choose a password"
            Char(Index).GameState = "PasswordChoosing"
        Case "n", "no"
            Send Index, "Alright, choose a new one then"
            Char(Index).GameState = "Name"
        Case Else
            Send Index, "Yes or no, please"
        End Select
    Case "PasswordChoosing"
        Char(Index).GameSubState = Data 'Storing the Password to GameSubState of char
        Send Index, "Retype the password again"
        Char(Index).GameState = "PasswordConfirm"
    Case "PasswordConfirm"
        'Checks the password and that the player isn't online
        If Char(Index).GameSubState = Data Then
            Send Index, "Password confirmed"
            Send Index, "Please choose a race: human/elf/gnome/kendar/dwarf"
            Char(Index).GameState = "Race"
        Else
            Send Index, "The passwords don't match. Please choose your password again."
            Char(Index).GameState = "PasswordChoosing"
        End If
    Case "Race"
        Select Case LowData
        Case "human", "gnome", "elf", "kendar", "dwarf"
            Char(Index).Race = LowData
            Send Index, "You are now " & bWHITE & LowData & WHITE
            Send Index, "I can't see very well here, are you male or female?"
            Char(Index).GameState = "Gender"
        Case Else
            Send Index, "This race must be from some other mud... Choose again: human/elf/gnome/kendar/dwarf"
        End Select
    Case "Gender"
        Select Case LowData
        Case "male", "m"
            Char(Index).Gender = "male"
            Char(Index).GameState = "newchar"
            CreationFinish Index
        Case "female", "f"
            Char(Index).Gender = "female"
            Char(Index).GameState = "newchar"
            CreationFinish Index
        Case Else
            Send Index, "Yet, pretend to be one of these: Male/Female"
        End Select

'\B/----------------------------Character already exists----------------------------
    Case "PasswordCheck"
    Dim I%
        If DbFind(CharRecord.SnapShot, CharRecord.Table, Index, "Name", Char(Index).Name, "Password") = Data Then
            For I = 1 To frmMain.wskAccept.UBound
                If Char(I).Name = Char(Index).Name And Char(I).GameState = "Game" Then _
                  Send Index, "This character is already PLAYING. That stinks =(.": Char(Index).GameState = "Name": Exit Sub
            Next I
            'Check is character is playing ALREADY
            Char(Index).GameState = "Game"
            CreationFinish Index
        Else
            Send Index, "Wrong password. Enter your name again.": Char(Index).GameState = "Name"
        End If
'/E\----------------------------Character already exists----------------------------
    End Select
End Sub
Sub CreationFinish(Index As Integer)
    
    If Char(Index).GameState = "newchar" Then
        Send Index, "Flesh grows in the right parts of your body as you choose a gender" & RET
        '> Saving to database
        CharRecord.Dynaset.AddNew
            CharRecord.Dynaset("Name") = Char(Index).Name
            CharRecord.Dynaset("Password") = Char(Index).GameSubState 'Contains password
            CharRecord.Dynaset("Race") = Char(Index).Race
            CharRecord.Dynaset("Sex") = Char(Index).Gender
            CharRecord.Dynaset("Phand") = "right"
            CharRecord.Dynaset("Status") = "mortal"
            CharRecord.Dynaset("BodyPartsCondition") = "100,100,100,100,100"
            CharRecord.Dynaset("BodyPartsAC") = "0,0,0,0,0"
        CharRecord.Dynaset.Update
        
        Set CharRecord.SnapShot = DB.OpenRecordset("Characters", dbOpenSnapshot)  'Making the characters' RecordSet
        CharRecord.SnapShot.MoveLast
    End If
    
    Char(Index).Area = 1
    Char(Index).GameState = "Game"
    Char(Index).locX = 1: Char(Index).locY = 1: Char(Index).locZ = 0
    Char(Index).Gender = CharRecord.SnapShot!Sex
    Char(Index).Race = CharRecord.SnapShot!Race
    Char(Index).Items = CharRecord.SnapShot!Items & ""
    Char(Index).PHand.Name = CharRecord.SnapShot!PHand
    Char(Index).Wear = CharRecord.SnapShot!Wear & ""
    Char(Index).Status = GDic(CStr(CharRecord.SnapShot!Status))
    
    Char(Index).HPMax = BodyMaxHP
    
    If InStr(Char(Index).Wear, "phand") = 0 Then
'\B/-------------If character doesn't hold any weapon, it uses bare hands-------------
    Char(Index).PHand.WearVnum = 1
    Char(Index).Damage = 1
'/E\-------------If character doesn't hold any weapon, it uses bare hands-------------
    Else
        Char(Index).Damage = Item(Char(Index).PHand.WearVnum).Damage
    End If
    
   Dim Cond$, AC$
    
    Cond = CharRecord.SnapShot!BodyPartsCondition
    AC = CharRecord.SnapShot!BodyPartsAC

    Char(Index).Head.Cond = GetWordByNum(1, Cond)
    Char(Index).Torso.Cond = GetWordByNum(2, Cond)
    Char(Index).Legs.Cond = GetWordByNum(3, Cond)
    Char(Index).PHand.Cond = GetWordByNum(4, Cond)
    Char(Index).SHand.Cond = GetWordByNum(5, Cond)
    
    Char(Index).Head.AC = GetWordByNum(1, AC)
    Char(Index).Torso.AC = GetWordByNum(2, AC)
    Char(Index).Legs.AC = GetWordByNum(3, AC)
    Char(Index).PHand.AC = GetWordByNum(4, AC)
    Char(Index).SHand.AC = GetWordByNum(5, AC)
    
    Char(Index).Head.Name = "head"
    Char(Index).Torso.Name = "torso"
    Char(Index).Legs.Name = "legs"
    Char(Index).SHand.Name = IIf(Char(Index).PHand.Name = "right", "left", "right") & " hand"
    Char(Index).PHand.Name = Char(Index).PHand.Name & " hand"
        
    Char(Index).HP = BodyMaxHP%
    
    Select Case Char(Index).Status
    Case Admin
        PlayerList.Admins = AddToString(PlayerList.Admins, Trim(Index))
    Case Immortal
        PlayerList.Immortals = AddToString(PlayerList.Immortals, Trim(Index))
    Case Mortal
        PlayerList.Mortals = AddToString(PlayerList.Mortals, Trim(Index))
    End Select
    AddPC Index
    Send Index, "Welcome to BeMUD, " & Char(Index).Name & vbCrLf  'Creation stage is over
    Send Index, "Type " & bWHITE & "HELP" & WHITE & " to get the list of the commands" & RET
    pcmdLook (Index)
    TransmitLocal Index, Char(Index).Name & " wakes up, yawning."
End Sub
