Attribute VB_Name = "mdlPc"
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
    '> Consts
        Public Const BodyMaxHP% = 100
        
    
    '< Consts
    '> Characters information variables
        Enum StatusEnum
            Admin = 3
            Immortal = 2
            Mortal = 1
        End Enum
        Type BodyPartVars
            Name As String 'Name of the part
            Cond As Integer '100 highest, 0 lowest
            AC As Integer '100 highest, 0 lowest
            WearVnum As Integer 'Vnum of item weared on this body part
        End Type
        Type DelayVars
            Busy As Boolean
            Command As String
            PCTarget As Integer
            MobItemVnum As Integer
         End Type
    Public Char() As Character
    Public Type Character
        Name As String 'Character's name
        Gender As String
        Race As String 'Character's race
        Area As Integer 'The ID of the area Char() is in
        locX As Integer ' \
        locY As Integer '  > Location on the current area (X,Y,Z)
        locZ As Integer ' /
        HP As Integer 'Blood in body
        HPMax As Integer
    '> Char body parts
        Head As BodyPartVars
        Torso As BodyPartVars
        Legs As BodyPartVars
        PHand As BodyPartVars
        SHand As BodyPartVars
    '< Char body parts
        Items As String 'Items the char has (Inventory)
        Wear As String 'All the items one char wears
        Damage As Integer 'Damage char does in a hit
        ApproachedPCs As String 'PCs that approached the char
        ApproachedMobs As String 'Mobs that approached the char
        GameState As String 'GameState defines if some commands will work or not
        GameSubState As String 'More common changing states
        Data As String 'That's each char unique input string
        Record As Recordset 'Thevar holds the Record movement table
        Bleeding As Integer 'The amount of blood you bleed every delayed.Bleeding
        Delay As DelayVars
        Status As StatusEnum 'Admin/Immortal/Mortal
        TimeOnline As Integer 'Time online (DUH)
        Spy As Boolean 'The admin command to spy after player/s
        QdText As String 'The text that doesn't fit to one screen of terminal
    End Type
    Public PlayerList As PlayerLists
    Type PlayerLists
        Admins As String
        Immortals As String
        Mortals As String
        
        Spy As String
        
        FreeIndex As String 'Stores the indexes of people that quited
    End Type
    
'< Characters information variables

Function OpenEmoteTags(ByVal Str$, Index%, Arguement$, Optional Targeted%) As String
'\B/------------------------------Open  PC related tags------------------------------
    If Targeted > 0 Then Str = Replace(Str, "<target>", Char(Targeted).Name): Arguement = Replace(LCase(Arguement), LCase(Char(Targeted).Name), "")
    Str = Trim(Replace(Str, "<arg>", Arguement))
    If InStr(Str, "<hisher>") Then Str = Replace(Str, "<hisher>", HeShe(Char(Index).Gender, "HisHer"))
    OpenEmoteTags = Str & "."
'/E\------------------------------Open  PC related tags------------------------------
End Function
Sub GetCharLook(ByVal NameIndex As Integer, ByRef Condition$, ByRef WearedItemsLook$)
    '\B/----------------------------------------See wearing equipment----------------------------------------
       WearedItemsLook = WearsItems(SortItemsList(Char(NameIndex).Wear), _
                                     Char(NameIndex).PHand.Name, Char(NameIndex).Gender)
    '/E\----------------------------------------See wearing equipment----------------------------------------
    '\B/--------------------------------------------See condition--------------------------------------------
    Dim HisHer$
        
        With Char(NameIndex)
        
            HisHer = HeShe(Char(NameIndex).Gender, "hisher")
            If .Head.Cond < BodyMaxHP% Then _
              Condition = Condition & HisHer & " head is " & GetCondition(.Head.Cond, .HPMax) & ", "
            If .Torso.Cond < BodyMaxHP% Then _
              Condition = Condition & HisHer & " torso is " & GetCondition(.Torso.Cond, .HPMax) & ", "
            If .PHand.Cond < BodyMaxHP% Then _
              Condition = Condition & HisHer & " " & .PHand.Name & " is " & GetCondition(.PHand.Cond, .HPMax) & ", "
            If .SHand.Cond < BodyMaxHP% Then _
              Condition = Condition & HisHer & " " & .SHand.Name & " is " & GetCondition(.SHand.Cond, .HPMax) & ", "
            If .Legs.Cond < BodyMaxHP% Then _
              Condition = Condition & HisHer & " legs are " & GetCondition(.Legs.Cond, .HPMax) & ", "
            If Condition <> "" Then Condition = Proper(Left(Condition, Len(Condition) - 2)) & "."
            
        End With
    '/E\--------------------------------------------See condition--------------------------------------------
    End Sub
Sub DoMobHit(ByVal Index%, Mnum%, ByRef HittedBodyPart As BodyPartVars)
    Dim Damage%, BleedMaxTemp%
    Damage = Char(Index).Damage
    If HittedBodyPart.Name = "head" Then If Int(Rnd * 2) + 1 = 2 Then Damage = Damage + 3
        
        If Damage <= HittedBodyPart.AC Then
            If HittedBodyPart.WearVnum > 0 Then Send Index, "Your blow falls on " & .Name & "'s " & Item(HittedBodyPart.WearVnum).Name & " but doesn get throu."
            If HittedBodyPart.WearVnum = 0 Then Send Index, "Your hit hardly touches " & .Name & " " & HittedBodyPart.Name & "'s skin."
            Exit Sub
        End If
        
End Sub
Function HeShe(Gender$, GenderType$) As String
    GenderType = LCase(GenderType)
    If Gender = "male" Then
        Select Case GenderType
        Case "heshe"
            HeShe = "he"
        Case "hisher"
            HeShe = "his"
        End Select
    Else
        Select Case GenderType
        Case "heshe"
            HeShe = "she"
        Case "hisher"
            HeShe = "her"
        End Select
    End If
End Function
Function AllUsers() As Variant
    AllUsers = AddToString(PlayerList.Mortals, AddToString(PlayerList.Admins, PlayerList.Immortals))
End Function
Function PcIsHere(Index As Integer, Name As String) As Integer
    'Returns the index of the found char (PcIsHere>0)
    Dim I As Integer
    Dim Arr As Variant
    PcIsHere = 0
    Arr = StringToArray(PCs(Index))
    If UBound(Arr) > 0 Then
        For I = LBound(Arr) To UBound(Arr)
            If Arr(I) <> Index And LCase(Char(Arr(I)).Name) = Name Then PcIsHere = Arr(I): Exit Function
        Next I
    End If
End Function
'Function that checks if what index of array is equal to LookFor
Function SearchCharName(ArrayName() As Character, LookFor As String) As Integer
    Dim Arr$(), I%
    SearchCharName = 0
    For I = LBound(ArrayName) To UBound(ArrayName)
        If ArrayName(I).Name = LookFor Then SearchCharName = I: Exit Function
    Next I
End Function
Sub ApproachRemoval(Index As Integer)
With Char(Index)

Dim Arr As Variant, I As Integer

'\B/------------------------If approached PCs then retreating------------------------
    If Len(.ApproachedPCs) > 0 Then
        Arr = StringToArray(.ApproachedPCs)
        For I = LBound(Arr) To UBound(Arr)
            Char(Arr(I)).ApproachedPCs = RemoveFromString(Char(Arr(I)).ApproachedPCs, Index)
        Next I
        .ApproachedPCs = ""
    End If
'/E\------------------------If approached PCs then retreating------------------------

'\B/-----------------------If approached mobs then retreating-----------------------
    If Len(.ApproachedMobs) > 0 Then
        Arr = StringToArray(.ApproachedMobs)
        For I = LBound(Arr) To UBound(Arr)
            Mob(Arr(I)).ApproachedPCs = _
              RemoveFromString(Mob(Arr(I)).ApproachedPCs, Index)
        Next I
        .ApproachedMobs = ""
    End If
'/E\-----------------------If approached mobs then retreating-----------------------

End With
End Sub
Function IsApproached(Index%) As Boolean
With Char(Index)

    If Len(.ApproachedPCs) + Len(.ApproachedMobs) > 0 Then _
      IsApproached = True Else IsApproached = False

End With
End Function
Function GetBodyPart(NameIndex%, MobItemVnum%) As BodyPartVars
If CheckMPI(NameIndex) Then
    With Char(NameIndex)
        'Where the blow falls
        Select Case Int(Rnd * 21) + 1
        Case 1 To 10
            GetBodyPart = .Torso
        Case 11 To 13
            GetBodyPart = .PHand
        Case 14 To 16
            GetBodyPart = .SHand
        Case 17 To 20
            GetBodyPart = .Legs
        Case 21
            GetBodyPart = .Head
        End Select
    End With
    Exit Function
End If
If CheckMPI(MobItemVnum) Then
    With Mob(MobItemVnum)
        'Where the blow falls
        Select Case Int(Rnd * 21) + 1
        Case 1 To 10
            GetBodyPart = .Torso
        Case 11 To 13
            GetBodyPart = .PHand
        Case 14 To 16
            GetBodyPart = .SHand
        Case 17 To 20
            GetBodyPart = .Legs
        Case 21
            GetBodyPart = .Head
        End Select
    End With
    Exit Function
End If
End Function
Sub LetBodyPart(NameIndex%, MobItemVnum%, HitPart As BodyPartVars)

If CheckMPI(NameIndex) Then
    With Char(NameIndex)
        
        Select Case HitPart.Name
        Case .Torso.Name
            .Torso = HitPart
        Case .PHand.Name
            .PHand = HitPart
        Case .SHand.Name
            .SHand = HitPart
        Case .Legs.Name
            .Legs = HitPart
        Case .Head.Name
            .Head = HitPart
        End Select
    
    End With
    Exit Sub
End If
If CheckMPI(MobItemVnum) Then
    With Mob(MobItemVnum)
        Select Case HitPart.Name
        Case .Torso.Name
            .Torso = HitPart
        Case .PHand.Name
            .PHand = HitPart
        Case .SHand.Name
            .SHand = HitPart
        Case .Legs.Name
            .Legs = HitPart
        Case .Head.Name
            .Head = HitPart
        End Select
    End With
    Exit Sub
End If

End Sub

Function ACDefended(Index%, NameIndex%, Damage As Integer, HittedBodyPart As BodyPartVars, _
                    Optional MissMsgSelf$ = "", Optional MissMsgTarget$ = "", Optional MissMsgOthers$ = "") As Boolean
    ACDefended = False
    If Damage <= HittedBodyPart.AC Then
        ACDefended = True
        Send Index, MissMsgSelf
        If NameIndex > 0 Then Send NameIndex, MissMsgTarget
        TransmitLocal Index, MissMsgOthers, NameIndex
    End If
End Function
Function DoDamage(NameIndex%, Damage%, HittedBodyPart As BodyPartVars) As Integer
    Dim BleedMaxTemp As Integer
    
    DoDamage = Damage
    
    With Char(NameIndex)
    HittedBodyPart.Cond = HittedBodyPart.Cond - Damage + HittedBodyPart.AC
    
    .Bleeding = Abs(.Torso.Cond + .Head.Cond _
      + .Legs.Cond + .PHand.Cond + .SHand.Cond - BodyMaxHP% * 5) / (BodyMaxHP% / 5)
    
    If .Bleeding > 0 And frmMain.tmrBleeding(NameIndex).Enabled = False _
      Then frmMain.tmrBleeding(NameIndex).Enabled = True
    
'\B/--------------------------Setting the bleeding interval--------------------------
    BleedMaxTemp = Int((BodyMaxHP% * 5 - .Bleeding * 5) / 25)
    frmMain.tmrBleeding(NameIndex).Interval = BleedMaxTemp * 1000
'/E\--------------------------Setting the bleeding interval--------------------------
    
'\B/---------------------------Lowering the overall health---------------------------
    .HP = .HP - Damage + HittedBodyPart.AC
'/E\---------------------------Lowering the overall health---------------------------
    
    End With
    
End Function
Function DoDamageMob(TargetMnum%, Damage%, HittedBodyPart As BodyPartVars) As Integer
    Dim BleedMaxTemp As Integer
    
    DoDamageMob = Damage
    
    With Mob(TargetMnum)
        
        HittedBodyPart.Cond = HittedBodyPart.Cond - Damage + HittedBodyPart.AC
        .HP = .HP - Damage + HittedBodyPart.AC
    
        .Bleeding = Abs(.Torso.Cond + .Head.Cond _
          + .Legs.Cond + .PHand.Cond + .SHand.Cond - .MaxHP * 5) / (.MaxHP / 5)
        
        If .Bleeding > 0 And frmMain.tmrBleedingMob(TargetMnum).Enabled = False _
          Then frmMain.tmrBleedingMob(TargetMnum).Enabled = True
    
'\B/--------------------------Setting the bleeding interval--------------------------
        BleedMaxTemp = Int((.MaxHP * 5 - Mob(TargetMnum).Bleeding * 5) / 25)
        frmMain.tmrBleedingMob(TargetMnum).Interval = BleedMaxTemp * 1000
'/E\--------------------------Setting the bleeding interval--------------------------
    
'\B/---------------------------Lowering the overall health---------------------------
        Mob(TargetMnum).HP = Mob(TargetMnum).HP - Damage + HittedBodyPart.AC
'/E\---------------------------Lowering the overall health---------------------------
    
    End With
End Function
Sub DelayedAttack(Index%, HittedBodyPart As BodyPartVars, _
             Optional MissMsgSelf$ = "", Optional MissMsgOthers$ = "", Optional MissMsgTarget$, _
             Optional MsgSelf$ = "", Optional MsgOthers$ = "", Optional MsgTarget$ = "")
    Dim NameIndex% 'Holds the found Index of a Name
    Dim TargetMnum As Integer 'Holds the found Vnum of the mob
    Dim DamageDone As Integer 'The damage that was done
    
    
    NameIndex = Char(Index).Delay.PCTarget
    TargetMnum = Char(Index).Delay.MobItemVnum
        
    If TargetChanged(Index, NameIndex, TargetMnum) Then Exit Sub 'Is the target here?
                
'\B/-----------------------------------PC to PC hit-------------------------------------
    If CheckMPI(NameIndex) Then
        
        With HittedBodyPart
            MissMsgSelf = GlobalOpenTags(MissMsgSelf, Index:=NameIndex)
            MissMsgTarget = GlobalOpenTags(MissMsgTarget, Index:=NameIndex)
            MissMsgOthers = GlobalOpenTags(MissMsgOthers, Index:=NameIndex)
            
            MsgSelf = GlobalOpenTags(MsgSelf, Index:=NameIndex, BodyPartName:=HittedBodyPart.Name)
            MsgTarget = GlobalOpenTags(MsgTarget, Index:=NameIndex, BodyPartName:=HittedBodyPart.Name)
            MsgOthers = GlobalOpenTags(MsgOthers, Index:=NameIndex, BodyPartName:=HittedBodyPart.Name)
        End With
        
        If ACDefended(Index, NameIndex, Char(Index).Damage, HittedBodyPart, MissMsgSelf, _
           MissMsgTarget, MissMsgOthers) Then Exit Sub
        
        DamageDone = DoDamage(NameIndex, Char(Index).Damage, HittedBodyPart)
    
    '\B/-----------------------------------Transmiting-----------------------------------
        Send Index, MsgSelf
        Send NameIndex, MsgTarget
        TransmitLocal Index, MsgOthers, NameIndex
    '/E\-----------------------------------Transmiting-----------------------------------
    End If
'/E\-----------------------------------PC TO PC hit-------------------------------------
    
    '\B/-----------------------------------PC to MOB hit-------------------------------------
    If CheckMPI(TargetMnum) Then
        
        With Mob(TargetMnum)
        MissMsgSelf = GlobalOpenTags(MissMsgSelf, Mnum:=TargetMnum, BodyPartName:=HittedBodyPart.Name)
        MissMsgTarget = GlobalOpenTags(MissMsgTarget, Mnum:=TargetMnum, BodyPartName:=HittedBodyPart.Name)
        MissMsgOthers = GlobalOpenTags(MissMsgOthers, Mnum:=TargetMnum, BodyPartName:=HittedBodyPart.Name)
        
        MsgSelf = GlobalOpenTags(MsgSelf, Mnum:=TargetMnum, BodyPartName:=HittedBodyPart.Name)
        MsgTarget = GlobalOpenTags(MsgTarget, Mnum:=TargetMnum, BodyPartName:=HittedBodyPart.Name)
        MsgOthers = GlobalOpenTags(MsgOthers, Mnum:=TargetMnum, BodyPartName:=HittedBodyPart.Name)
        End With
        
        If ACDefended(Index, NameIndex, Char(Index).Damage, HittedBodyPart, MissMsgSelf, _
           MissMsgTarget, MissMsgOthers) Then Exit Sub
           
        Call DoDamageMob(TargetMnum, Char(Index).Damage, HittedBodyPart)
        
'\B/-------------------------------------Mob AI-------------------------------------
        Call AIGotHitMob(TargetMnum, DamageDone, Index)
'/E\-------------------------------------Mob AI-------------------------------------
        
        Send Index, MsgSelf
        TransmitLocal Index, MsgOthers
        
    End If
    '/E\-----------------------------------PC TO MOB hit-------------------------------------

    'Updating the body part
    With Char(Index).Delay
        Call LetBodyPart(.PCTarget, .MobItemVnum, HittedBodyPart)
    End With
    
End Sub
Sub Attack(Index%, Arguement$, NotHere$, Busy$, TooFar$, _
    PreparingSelf$, PreparingTarget$, PreparingOthers$, DelayCommand$, Seconds%, _
    Optional CheckForApproach As Boolean = True)

    Dim NameIndex%, Mnum As Integer
    
    If Arguement = "" Then Send Index, "What?": Exit Sub
    
    NameIndex = PcIsHere(Index, LCase(Arguement))
    Mnum = MobIsHere(Index, LCase(Arguement))
    
    'Checking if the char can hit at all
    If Not CheckMPI(NameIndex) And Not CheckMPI(Mnum) Then Send Index, NotHere: Exit Sub
    
'\B/--------------------------------Checking if busy--------------------------------
    With Char(Index).Delay
    If .Busy = True Then Send Index, Busy: Exit Sub
    End With
'/E\--------------------------------Checking if busy--------------------------------
    
    If CheckMPI(NameIndex) Or CheckMPI(Mnum) Then
        
        If CheckMPI(NameIndex) Then
            
            With Char(NameIndex)
            PreparingSelf = GlobalOpenTags(PreparingSelf, Index:=NameIndex)
            PreparingTarget = GlobalOpenTags(PreparingTarget, Index:=NameIndex)
            PreparingOthers = GlobalOpenTags(PreparingOthers, Index:=NameIndex)
            End With
            
            If CheckForApproach Then _
             If InStr(QteMe(Char(Index).ApproachedPCs), QteMe(NameIndex)) = 0 Then Send Index, TooFar: Exit Sub
            
            Send Index, PreparingSelf
            Send NameIndex, PreparingTarget
            TransmitLocal Index, PreparingOthers, NameIndex
        
        ElseIf CheckMPI(Mnum) Then
        
            With Mob(Mnum)
            PreparingSelf = GlobalOpenTags(PreparingSelf, Mnum:=Mnum)
            PreparingOthers = GlobalOpenTags(PreparingOthers, Mnum:=Mnum)
            End With
         
            If CheckForApproach Then _
            If InStr(QteMe(Char(Index).ApproachedMobs), QteMe(Mnum)) = 0 Then Send Index, TooFar: Exit Sub
            
            Send Index, PreparingSelf
            TransmitLocal Index, PreparingOthers
        
        End If
    
        Call AddDelay(Index, Seconds:=Seconds, CommandName:=DelayCommand, Mnum:=Mnum, NameIndex:=NameIndex)
    End If

End Sub
