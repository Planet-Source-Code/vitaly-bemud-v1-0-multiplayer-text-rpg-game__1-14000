Attribute VB_Name = "mdlPcCommands"
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
Sub DoCommands(Index As Integer, ByVal Data$)
    Dim Command As String
    Dim Arguement As String
    Dim I As Integer
    Dim NameIndex% 'Holds the found Index of a Name
    Dim Mnum 'Holds the vbyn if tge mob
    Dim Inum 'Holds the vnum of the item
    Dim ItemsValue As Variant 'Holds the sorted by order ItemsValue to display.
'> Resolves the command from everything else
    Char(Index).Data = ""
    If InStr(Data, " ") Then
        Command = Left(Data, InStr(Data, " ") - 1)
        Arguement = Mid(Data, InStr(Data, " ") + 1)
    Else
        Command = Data
    End If
    If Left(Data, 1) = "'" Or Left(Data, 1) = ";" Then
        Command = Left(Data, 1)
        Arguement = Right(Data, Len(Data) - 1)
    End If
'> Resolves the command from everything else
'> Expanding aliases
        Select Case Command
        Case "n"
            Command = "north"
        Case "s"
            Command = "south"
        Case "w"
            Command = "west"
        Case "e"
            Command = "east"
        Case "u"
            Command = "up"
        Case "d"
            Command = "down"
        End Select
'< Expanding aliased movements
'\B/--------------------------------Movement commands--------------------------------
    If InStr(QteMe(CurrentExits(Index)), QteMe(Command)) Then
        Call pcmdMovement(Index, Command)
        Exit Sub
    End If
'/E\--------------------------------Movement commands--------------------------------
    Dim EmoteIndex%
    EmoteIndex = SearchEmoteID(Emotes(), Command) 'Assigns EmoteIndex WhatMember findings
'> Other commands
    Select Case Command
'******************************************* A *****************************************
    Case "approach"
        Call pcmdApproach(Index, Command, Arguement)
'******************************************* B *****************************************
'******************************************* C *****************************************
'******************************************* D *****************************************
'COMMAND: Drop
    Case "drop"
        Call pcmdDrop(Index, Command, Arguement)
'******************************************* E *****************************************
'COMMAND: "Emote"
    Case "emote", ";"
        Call pcmdEmote(Index, Command, Arguement)
'COMMANDS: Ready emotes
    Case IIf(CheckMPI(EmoteIndex), Command, Command & "!") 'Command! doesn't enter the case
        Call pcmdEmoteReady(Index, Command, Arguement, EmoteIndex)
'******************************************* F *****************************************
'******************************************* G *****************************************
'COMMAND: Gear
    Case "gear", "eq", "equipment"
        Call pcmdWearShow(Index, Command, Arguement)
'COMMAND: Get
    Case "get"
        Call pcmdGet(Index, Command, Arguement)
    Case "give"
        Call pcmdGive(Index, Command, Arguement)
'******************************************* H *****************************************
'COMMAND: "Help"
    Case "help"
        Call pcmdHelp(Index, Command, Arguement)
        'COMMAND: "Hit"
    Case "hit"
        Call pcmdHit(Index, Command, Arguement)
'******************************************* I *****************************************
'COMMAND: Inventory
    Case "i", "inv", "inventory"
        Call pcmdInventory(Index, Command, Arguement)
'******************************************* J *****************************************
'******************************************* K *****************************************
'******************************************* L *****************************************
'COMMAND: LoadMob
    Case "loadmob"
        If Char(Index).Status = Admin Then Call pcmdLoadMob(Index, Command, Arguement)
'COMMAND: "Look"
    Case "look", "l"
        If Arguement = "" Then Call pcmdLook(Index) _
          Else: Call pcmdLookAt(Index, Command, Arguement) 'Look AT something
'******************************************* M *****************************************
    Case "missile"
        Call pcmdMissile(Index, Command, Arguement)
'******************************************* N *****************************************
'COMMAND: "News" for the latest code tweaks
    Case "news"
        Call pcmdNews(Index, Command, Arguement) 'Look AT something
'******************************************* O *****************************************
'******************************************* P *****************************************
'******************************************* Q *****************************************
' COMMAND: "Quit"
    Case "quit"
        Call pcmdQuit(Index, Command, Arguement) 'Look AT something
'******************************************* R *****************************************
'COMMAND: Remove - Remove equipped items
    Case "remove"
        Call pcmdRemove(Index, Command, Arguement) 'Look AT something
    Case "retreat"
        Call pcmdRetreat(Index, Command)
'******************************************* S *****************************************
'COMMAND: "Say"
    Case "say", "'"
        Call pcmdSay(Index, Command, Arguement) 'Look AT something

'COMMAND: Spy!
    Case "spy"
        If Char(Index).Status = Admin Then Call pcmdSpy(Index, Command, Arguement)
'******************************************* T *****************************************
'******************************************* U *****************************************
'******************************************* V *****************************************
'******************************************* W *****************************************
'COMMAND: Wear - Wear items (not weapons)
    Case "wear", "hold", "wield"
        If Arguement <> "" Then Call pcmdWear(Index, Command, Arguement) _
          Else: Call pcmdWearShow(Index, Command, Arguement)
    Case "who"
          Call pcmdWho(Index, Command)
'******************************************* X *****************************************
    'Case "xp"
        'DbEdit Index, "Characters", "Name", Char(Index).Name, dbCHARSTATUS, Int(Rnd * 10) + 1
'******************************************* Y *****************************************
'******************************************* Z *****************************************
    Case Else
        'If nothing works:
        Send Index, "What?"
        Debug.Print Command
    End Select
'< Other commands
End Sub
Sub pcmdMissile(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
Dim NotHere$, Busy$, TooFar$, DelayCommand$
Dim PreparingSelf$, PreparingTarget$, PreparingOthers$

NotHere = "It is not here."
Busy = "You're too busy to do that."
TooFar = "You're too far, approach first."

With Char(Index)
PreparingSelf = "You murmur to yourself a few MAGIC words."
PreparingTarget = .Name & " starts to murmur some mystic words."
PreparingOthers = .Name & " starts to murmur some mystic words."
End With

Call Attack(Index, Arguement, NotHere, Busy, TooFar, _
                   PreparingSelf, PreparingTarget, PreparingOthers, "missile-d", Seconds:=2)

End Sub
Sub pcmdMissileDelay(Index%)
    Dim MissMsgSelf$, MissMsgOthers$, MissMsgTarget$, MsgSelf$, MsgOthers$, MsgTarget$
    Dim HitBodyPart As BodyPartVars
    
    With Char(Index).Delay
        HitBodyPart = GetBodyPart(.PCTarget, .MobItemVnum)
    End With
        
    With Char(Index)
    
    MissMsgSelf = "The missles buzzz as they miss <targetname> by " & Int(Rnd * 5) + 1 & " meteres."
    MissMsgTarget = .Name & "'s V-2 missles buzzz as they miss you by a few meteres."
    MissMsgOthers = .Name & "'s V-2 missles buzzz as they miss <targetname> by a few meteres."
    
    MsgSelf = bMAGNETA & "Your pack of V-2 missiles fly and explode on <targetname>!" & WHITE
    MsgTarget = bMAGNETA & .Name & " sends V-2 missles and they explode on you!" & WHITE
    MsgOthers = bMAGNETA & .Name & " sends V-2 missles and they explode on <targetname>!" & WHITE
    
    End With
    
    Call DelayedAttack(Index, HitBodyPart, _
                MissMsgSelf, MissMsgOthers, MissMsgTarget, _
                MsgSelf, MsgOthers, MsgTarget)
End Sub
Sub pcmdMovement(Index As Integer, Command As String)
    Dim XYZ$
    
    If IsApproached(Index) Then Send Index, "You have to retreat first.": Exit Sub
    
    TransmitLocal Index, Char(Index).Name & " leaves " & Command & "."
    RemovePC Index
    XYZ = GetWordByNum(GetNumByWord(Command, CurrentExits(Index)) + 1, CurrentExits(Index))
    Char(Index).locX = GetWordByNum(1, XYZ, " ")
    Char(Index).locY = GetWordByNum(2, XYZ, " ")
    Char(Index).locZ = GetWordByNum(3, XYZ, " ")
    If CountWords(XYZ) > 3 Then Char(Index).Area = _
      SearchAreaName(Area(), GetWordByNum(4, XYZ, " ")) 'Link to other area
    AddPC Index
    TransmitLocal Index, Char(Index).Name & " arrives."
    pcmdLook (Index)
End Sub
Sub pcmdApproach(Index As Integer, Command As String, Arguement As String)
    Dim NameIndex%, Mnum As Integer
    Dim NotHere$, Busy$, TooFar$, DelayCommand$
    Dim PreparingSelf$, PreparingTarget$, PreparingOthers$
    
    If Arguement = "" Then Exit Sub
    
    NameIndex = PcIsHere(Index, LCase(Arguement))
    Mnum = MobIsHere(Index, LCase(Arguement))
    
    If Len(Char(Index).ApproachedPCs) > 0 Then
        If CheckMPI(NameIndex) Then If (InStr(QteMe(Char(Index).ApproachedPCs), QteMe(NameIndex)) > 0) _
           Then Send Index, "You're already standing there!": Exit Sub
        Send Index, "You have to retreat first.": Exit Sub
    End If
    
    If Len(Char(Index).ApproachedMobs) > 0 Then
        If CheckMPI(Mnum) Then If InStr(QteMe(Char(Index).ApproachedMobs), QteMe(Mnum)) > 0 _
           Then Send Index, "You're already standing there!": Exit Sub
        Send Index, "You have to retreat first.": Exit Sub
    End If
    
    NotHere = "It is not here."
    Busy = "You're too busy to do that."
    TooFar = ""
    
    With Char(Index)
    PreparingSelf = "You begin to approach <targetname>."
    PreparingTarget = .Name & " start to approach you!"
    PreparingOthers = .Name & " starts to approach <targetname>."
    End With

    Call Attack(Index, Arguement, NotHere, Busy, TooFar, _
                   PreparingSelf, PreparingTarget, PreparingOthers, "approach-d", _
                   Seconds:=2, CheckForApproach:=False)
     
End Sub
Sub pcmdApproachDelay(Index%)
    Dim NameIndex% 'Holds the found Index of a Name
    Dim Mnum As Integer 'Holds the vnum of the mob
    Dim Inum As Variant 'Holds the vnum of the item
     
    NameIndex = Char(Index).Delay.PCTarget
    Mnum = Char(Index).Delay.MobItemVnum
            
    If TargetChanged(Index, NameIndex, Mnum) Then Exit Sub 'Is the target here?
    
    If CheckMPI(NameIndex) Then
        Char(Index).ApproachedPCs = _
          AddToString(Char(Index).ApproachedPCs, Trim(NameIndex))
        Char(NameIndex).ApproachedPCs = _
          AddToString(Char(NameIndex).ApproachedPCs, Trim(Index))
        Send Index, "You come near " & Char(NameIndex).Name & "."
        Send NameIndex, Char(Index).Name & " comes and stands near you."
        TransmitLocal Index, Char(Index).Name & "comes and stands near " & Char(NameIndex).Name & ".", NameIndex
    End If
    
    If CheckMPI(Mnum) Then
        Char(Index).ApproachedMobs = AddToString(Char(Index).ApproachedMobs, Mnum)
        Mob(Mnum).ApproachedPCs = _
          AddToString(Mob(Mnum).ApproachedPCs, Trim(Index))
        Send Index, "You come near " & Mob(Mnum).Name & "."
        TransmitLocal Index, Char(Index).Name & " comes and stands near " & Mob(Mnum).Name & "."
    End If
End Sub
Sub pcmdLoadMob(Index As Integer, Command As String, Arguement As String)
    Dim Vnum%
    If IsNumeric(Arguement) Then
        With Char(Index)
            Vnum = LoadMob(.locX, .locY, .locZ, .Area, Index, Val(Arguement))
        End With
        If Vnum > 0 Then Send Index, PrototypeMob(Val(Arguement)).Name & " was created. (Vnum " & Vnum & ")" Else Send Index, "Invalid ID"
    Else
        Send Index, "Valid ID parameter is missing." & RET & "Syntax: LoadMob ID"
    End If
End Sub
Sub pcmdDrop(Index As Integer, Command As String, Arguement As String)
Dim Vnum As Integer
    Vnum = ItemIsInInv(Index, LCase(Arguement))
    If CheckMPI(Vnum) Then
        Call RemoveItemInv(Index, Val(Vnum))
        Call AddItem(Index, Val(Vnum))
        Send Index, "You drop " & Item(Vnum).Name & "."
        TransmitLocal Index, Char(Index).Name & " drops " & Item(Vnum).Name & "."
    Else
        Send Index, "You don't have it."
    End If
End Sub
Sub pcmdEmote(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
        Send Index, "-> " & Char(Index).Name & " " & Arguement
        TransmitLocal Index, "-> " & Char(Index).Name & " " & Arguement
End Sub
Sub pcmdEmoteReady(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String, EmoteIndex%)
Dim ToSelf$, ToOthers$, ToTarget$
Dim TargetIndex%, Mnum As Integer

    TargetIndex = PcIsHere(Index, LCase(GetWordByNum(1, Arguement)))
    Mnum = MobIsHere(Index, LCase(GetWordByNum(1, Arguement)))

    If TargetIndex > 0 Then
'\B/--------------------------------The target is PC--------------------------------
        ToSelf = "You " & Emotes(EmoteIndex).SelfTarget
        ToSelf = OpenEmoteTags(ToSelf, Index, StrConv(Arguement, vbProperCase), TargetIndex)
        ToTarget = Char(Index).Name & " " & Emotes(EmoteIndex).Target
        ToTarget = OpenEmoteTags(ToTarget, Index, Replace(Arguement, Char(TargetIndex).Name, ""), TargetIndex)
        ToOthers = Char(Index).Name & " " & Emotes(EmoteIndex).OthersTarget
        ToOthers = OpenEmoteTags(ToOthers, Index, StrConv(Arguement, vbProperCase), TargetIndex)
        Send Index, ToSelf
        Send PcIsHere(Index, LCase(GetWordByNum(1, Arguement))), ToTarget
        TransmitLocal Index, ToOthers, TargetIndex
'/E\--------------------------------The target is PC--------------------------------
    ElseIf CheckMPI(Mnum) Then
'\B/--------------------------------The target is mob--------------------------------
        ToSelf = "You " & Emotes(EmoteIndex).SelfTarget
        ToSelf = OpenMobEmoteTags(ToSelf, Index, StrConv(Arguement, vbProperCase), Mnum)
        ToOthers = Char(Index).Name & " " & Emotes(EmoteIndex).OthersTarget
        ToOthers = OpenMobEmoteTags(ToOthers, Index, StrConv(Arguement, vbProperCase), Mnum)
        Send Index, ToSelf
        TransmitLocal Index, ToOthers, TargetIndex
'/E\--------------------------------The target is mob--------------------------------
    Else
'\B/-----------------------------------------There is no target-----------------------------------------
        ToSelf = "You " & Emotes(EmoteIndex).Self
        ToSelf = OpenEmoteTags(ToSelf, Index, Arguement)
        ToOthers = Char(Index).Name & " " & Emotes(EmoteIndex).Others
        ToOthers = OpenEmoteTags(ToOthers, Index, Arguement)
        Send Index, ToSelf
        TransmitLocal Index, ToOthers
'/E\-----------------------------------------There is no target-----------------------------------------
    End If
End Sub
Sub pcmdGet(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
Dim Vnum As Integer
    Vnum = ItemIsHere(Index, LCase(Arguement))
    If CheckMPI(Vnum) Then
        Call RemoveItem(Index, Val(Vnum))
        Call AddItemInv(Index, Val(Vnum))
        Send Index, "You get " & Item(Vnum).Name & "."
        TransmitLocal Index, Char(Index).Name & " gets " & Item(Vnum).Name & "."
    Else
        Send Index, "It is not here."
    End If
End Sub
Sub pcmdGive(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
Dim NameIndex%, Inum As Integer
    If InStr(Arguement, " to ") Then
        NameIndex = PcIsHere(Index, LCase(Mid(Arguement, InStr(Arguement, " to ") + 4)))
        Inum = ItemIsInInv(Index, LCase(Mid(Arguement, 1, InStr(Arguement, " to ") - 1)))
        If CheckMPI(Inum) And CheckMPI(NameIndex) Then
            Call RemoveItemInv(Index, Val(Inum))
            Call AddItemInv(NameIndex, Val(Inum))
            Send Index, "You give " & Item(Inum).Name & " to " & Char(NameIndex).Name
            Send NameIndex, Char(Index).Name & " gives you " & Item(Inum).Name
            TransmitLocal Index, Char(Index).Name & " gives " & Item(Inum).Name & " to " & Char(NameIndex).Name, NameIndex
        End If
    Else
        Send Index, "Give command syntax: Give <what> to <whom>"
    End If
End Sub
Sub pcmdHelp(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
    Send Index, "Help is currently not working, see NEWS for the last changes"
End Sub
Sub pcmdHit(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
Dim NotHere$, Busy$, TooFar$, DelayCommand$
Dim PreparingSelf$, PreparingTarget$, PreparingOthers$

NotHere = "It is not here."
Busy = "You're too busy to do that."
TooFar = "You're too far, approach first."

With Char(Index)
PreparingSelf = "You aim to hit <targetname>."
PreparingTarget = .Name & " aims to hit you!"
PreparingOthers = .Name & " aims to hit <targetname>."
End With

Call Attack(Index, Arguement, NotHere, Busy, TooFar, _
                   PreparingSelf, PreparingTarget, PreparingOthers, "hit-d", Seconds:=2)

End Sub
Sub pcmdHitDelay(Index%)
    Dim MissMsgSelf$, MissMsgOthers$, MissMsgTarget$, MsgSelf$, MsgOthers$, MsgTarget$
    Dim HitBodyPart As BodyPartVars
    
    With Char(Index).Delay
        HitBodyPart = GetBodyPart(.PCTarget, .MobItemVnum)
    End With
        
    MissMsgSelf = "Your hit misses <targetname>."
    MissMsgTarget = Char(Index).Name & " misses you."
    MissMsgOthers = Char(Index).Name & " misses <targetname>."
    
    With Char(Index)
    MsgSelf = "You hit <targetname> on <bodypartname>!"
    MsgTarget = .Name & " hits you on <bodypartname>!"
    MsgOthers = .Name & " hits <targetname> on <bodypartname>!"
    End With
    
    Call DelayedAttack(Index, HitBodyPart, _
                MissMsgSelf, MissMsgOthers, MissMsgTarget, _
                MsgSelf, MsgOthers, MsgTarget)
End Sub
Sub pcmdInventory(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
    Dim WhatInInventory As String, Inventory As Variant
    Dim I%
    Inventory = StringToArray(Char(Index).Items)
    For I = 0 To UBound(Inventory)
        WhatInInventory = WhatInInventory & "   " & _
          Item(Inventory(I)).Name & _
          vbCrLf
    Next I
    Send Index, "You carry: " & vbCrLf & WhatInInventory
End Sub
Sub pcmdLook(ByVal Index As Integer)
'The look command
    Dim I%, ExitsHere$, WordsInCurrentExits%
    Dim WhoHere As Variant, ItemsHere As Variant
    Dim FullLook$ 'The var that holds all the look
    FullLook = FullLook & CurrentDesc(Index) & vbCrLf & vbCrLf 'Storing description of them room
'\B/----------------------------------Exits prompt--------------------------------
    ExitsHere = "You can go "
    WordsInCurrentExits = CountWords(CurrentExits(Index), ",")
    If WordsInCurrentExits Then
        For I = 1 To WordsInCurrentExits Step 2
            Select Case I
            Case 1
                ExitsHere = ExitsHere & GetWordByNum(I, CurrentExits(Index))
            Case 2 To WordsInCurrentExits - 2
                ExitsHere = ExitsHere & ", " & GetWordByNum(I, CurrentExits(Index))
            Case WordsInCurrentExits - 1
                ExitsHere = ExitsHere & " or " & GetWordByNum(I, CurrentExits(Index))
            End Select
        Next I
    Else
    ExitsHere = ExitsHere & "nowhere"
    End If
    ExitsHere = ExitsHere & "."
    FullLook = FullLook & bWHITE & ExitsHere & WHITE & vbCrLf
'/E\----------------------------------Exits prompt--------------------------------
'\B/----------------------------------------Represents other PCs----------------------------------------
    WhoHere = StringToArray(PCs(Index:=Index))
    If UBound(WhoHere) > 0 Then
        For I = LBound(WhoHere) To UBound(WhoHere)
            If WhoHere(I) <> Index Then FullLook = FullLook & Char(WhoHere(I)).Name & " is here." & RET
        Next I
        FullLook = FullLook & vbCrLf
    End If
'/E\----------------------------------------Represents other PCs----------------------------------------
'\B/-----------------------------------------Shows mobs in room-----------------------------------------
    WhoHere = StringToArray(Mobs(Index))
    If UBound(WhoHere) >= 0 Then
        For I = LBound(WhoHere) To UBound(WhoHere)
            FullLook = FullLook & Mob(Val(WhoHere(I))).Name & " is here." & RET
        Next I
        'FullLook = FullLook & vbCrLf
    End If
'/E\-----------------------------------------Shows mobs in room-----------------------------------------
'\B/-----------------------------------------Shows items in room-----------------------------------------
    ItemsHere = StringToArray(Items(Index))
    If UBound(ItemsHere) >= 0 Then
        For I = LBound(ItemsHere) To UBound(ItemsHere)
             FullLook = FullLook & Item(ItemsHere(I)).Name & RET
        Next I
    End If
'/E\-----------------------------------------Shows items in room-----------------------------------------
'Send it all to Index
    Send Index, FullLook
End Sub
Sub pcmdLookAt(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
    Dim Condition$, WearsItems$
    Dim NameIndex%, Mnum As Integer, Inum As Integer
'Looking for PCs
    NameIndex = PcIsHere(Index, LCase(Arguement))
'Looking for mobs
    Mnum = MobIsHere(Index, LCase(Arguement))
'Looking for items in inventory
    Inum = ItemIsInInv(Index, LCase(Arguement))
'If not found, looking in room
    If Not CheckMPI(Inum) Then Inum = ItemIsHere(Index, LCase(Arguement))
    If CheckMPI(NameIndex) Then
'\B/----------------------------------Looking at PC----------------------------------
        'Getting the information
        Call GetCharLook(NameIndex, Condition, WearsItems)
        
        Send Index, StrConv(HeShe(Char(NameIndex).Gender, "HeShe"), vbProperCase) & " is a " & _
          Char(NameIndex).Gender & " " & Char(NameIndex).Race & "." _
          & IIf(WearsItems <> "", vbCrLf, "") & WearsItems & IIf(Condition <> "", vbCrLf, "") & Condition
        Send NameIndex, Char(Index).Name & " looks at you."
        TransmitLocal Index, Char(Index).Name & " looks at " & Char(NameIndex).Name, NameIndex
'/E\----------------------------------Looking at PC----------------------------------
    ElseIf CheckMPI(Mnum) Then
'\B/---------------------------------Looking at mob---------------------------------
        Call GetMobLook(Mnum, Condition, WearsItems)
        With Mob(Mnum)
        
        Send Index, .Description _
          & IIf(WearsItems <> "", vbCrLf, "") & WearsItems & IIf(Condition <> "", vbCrLf, "") & Condition
        TransmitLocal Index, Char(Index).Name & " looks at " & .Name

        End With
'/E\---------------------------------Looking at mob---------------------------------
    ElseIf CheckMPI(Inum) Then
'\B/----------------------------------Look at item----------------------------------
        Send Index, Item(Inum).Description
        TransmitLocal Index, Char(Index).Name & " looks at " & Item(Inum).Name
'/E\----------------------------------Look at item----------------------------------
    Else
        Send Index, "It is not here"
    End If
End Sub
Sub pcmdNews(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
    Dim TempLine$, SumNews$
    On Error GoTo Error
    Dim News As String
    Open App.Path & "/" & "Tasks.dat" For Input As #1
    Do While Not EOF(1)
        Line Input #1, TempLine
        News = TempLine
        If InStr(News, "=") And Left(News, 4) = "Task" Then
            News = GetWordByNum(1, Mid(News, InStr(News, "=") + 1, Len(News)), "|")
            If Left(News, 1) = "*" Then News = Replace(News, "*", "Done: "): _
              News = News & vbCrLf & "      " & GetWordByNum(4, TempLine, "|") Else News = bWHITE & "Todo: " & WHITE & News
            SumNews = SumNews & News & vbCrLf
        End If
    Loop
    Send Index, SumNews, CheckForSplit:=True, Indent:="      "
Error:  Close #1
    If Err.Number > 0 Then Debug.Print Err.Description: Resume
End Sub
Sub pcmdQuit(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
    If Arguement <> "" Then Send Index, "To quit just type 'quit'.": Exit Sub
    If IsApproached(Index) Then Send Index, "You have to retreat first.": Exit Sub
    
    Send Index, "See you later ;)"
    Call CloseConnection(Index)
End Sub
Sub pcmdRemove(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
Dim Vnum As Integer
    Vnum = ItemIsWeared(Index, LCase(Arguement))
    If CheckMPI(Vnum) Then
        Call RemoveWearItem(Index, Val(Vnum))
        Call AddItemInv(Index, Val(Vnum))
        Send Index, "You remove " & Item(Vnum).Name
        TransmitLocal Index, Char(Index).Name & " removes " & Item(Vnum).Name
    Else
        Send Index, "You don't wear it."
    End If
End Sub
Sub pcmdRetreat(ByVal Index As Integer, ByVal Command As String)
    Dim NameIndex%, Mnum As Integer
    
    If Char(Index).Delay.Busy Then Send Index, "You're too busy": Exit Sub
    
    If Len(Char(Index).ApproachedPCs) > 0 Or Len(Char(Index).ApproachedMobs) > 0 Then
    Send Index, "You begin to retreat."
    Call AddDelay(Index, Seconds:=2, CommandName:="retreat-d", Mnum:=Mnum, NameIndex:=NameIndex)
    End If
End Sub
Sub pcmdRetreatDelay(ByVal Index As Integer)
    Dim Arr, I%
    Call ApproachRemoval(Index) 'Removes the char Index from other characters approached records
    Send Index, "You retreat."
    TransmitLocal Index, Char(Index).Name & " retreats."
End Sub
Sub pcmdSay(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
    Dim Temp$
    Temp = Timer
    If Arguement <> "" Then
        Send Index, "You say, '" & Arguement & "'"
        TransmitLocal Index, Char(Index).Name & " says, '" & Arguement & "'"
    End If
End Sub
Sub pcmdSpy(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
    Char(Index).Spy = Not Char(Index).Spy
    If Char(Index).Spy Then
        Send Index, "Spy function is ON."
        PlayerList.Spy = AddToString(PlayerList.Spy, Index)
    Else
        Send Index, "Spy function is OFF."
        PlayerList.Spy = RemoveFromString(PlayerList.Spy, Index)
    End If
End Sub
Sub pcmdWear(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
Dim Vnum As Integer
'To wear/hold an item
    
    Vnum = ItemIsInInv(Index, LCase(Arguement))
    If CheckMPI(Vnum) Then 'Checks if the item was found at all

'\B/----------------Checking if command is for the right weapon type----------------
        If ((Command = "wield" Or Command = "hold") And Item(Vnum).Type = "weapon") _
        Or (Command = "wear" And Item(Vnum).Type = "armor") Then
'/E\----------------Checking if command is for the right weapon type----------------
            If InStr(Char(Index).Wear, Item(Vnum).Wear) = 0 Then
                    Call RemoveItemInv(Index, Val(Vnum))
                    'Saves the weapon stats
                    Call WearItem(Index, Val(Vnum))
                    Send Index, "You " & Command & " " & Item(Vnum).Name
                    TransmitLocal Index, Char(Index).Name & " " & Command & "s " & Item(Vnum).Name & "."
            Else
                    Send Index, "You're already wearing something there."
            End If
        Else
            Send Index, "You can't do it."
        End If
    Else
        Send Index, "You don't have it."
        End If
End Sub

Sub pcmdWearShow(ByVal Index As Integer, ByVal Command As String, ByVal Arguement As String)
Dim Vnum As Integer, ItemsValue As Variant
Dim I%
'\B/-------------------------------Check what you wear-------------------------------
        Dim EquipmentOutput$
        ItemsValue = SortItemsList(Char(Index).Wear)
        For I = 1 To 5
            If ItemsValue(I) <> "" Then
                Select Case I
                Case 1
                    EquipmentOutput$ = _
                    StrConv(Char(Index).PHand.Name, vbProperCase) & ": " & _
                    Item(ItemsValue(I)).Name & vbCrLf
                Case 2
                    EquipmentOutput$ = EquipmentOutput$ & _
                      "Torso: " & Item(ItemsValue(I)).Name & vbCrLf
                Case 3
                    EquipmentOutput$ = EquipmentOutput$ & _
                      "Hands: " & Item(ItemsValue(I)).Name & vbCrLf
                Case 3
                    EquipmentOutput$ = EquipmentOutput$ & _
                      "Legs: " & Item(ItemsValue(I)).Name & vbCrLf
                End Select
            End If
        Next I
        If EquipmentOutput = "" Then EquipmentOutput = "You don't wear anything." & vbNewLine
        Send Index, EquipmentOutput, , , ""
'/E\-------------------------------Check what you wear-------------------------------
End Sub
Sub pcmdWho(ByVal Index As Integer, ByVal Command As String)
    Dim Arr, I%, Message$
    Arr = StringToArray(AllUsers)
'\B/----------------------------Compiling the WHO message----------------------------
    Message = "There " & IIf(UBound(Arr) + 1 > 1, "are", "is") & " currently " & UBound(Arr) + 1 & " player" & IIf(UBound(Arr) + 1 > 1, "s", "") & " in BeMUD." & RET
    For I = LBound(Arr) To UBound(Arr)
        Message = Message & "  " & Char(Arr(I)).Name & RET
    Next I
'/E\----------------------------Compiling the WHO message----------------------------
    Send Index, Message
End Sub
