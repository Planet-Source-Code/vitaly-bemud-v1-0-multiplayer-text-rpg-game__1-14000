Attribute VB_Name = "mdlMob"
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
'> Mobile vars
        Type AIVars
        
'\B/------------------------Mob character attributes 1 to 100------------------------
            Aggresion As Integer
            Dextrity As Integer
'/E\------------------------Mob character attributes 1 to 100------------------------
        
        End Type
    Public Type MobVars
        ID As Integer
        Name As String
        Aliases As String
        Description As String
        Gender As String
        Type As String
        Subtype As String
        HP As Integer
        MaxHP As Integer

'\B/---------------------------------Mob body parts----------------------------------
        Head As BodyPartVars
        Torso As BodyPartVars
        Legs As BodyPartVars
        PHand As BodyPartVars
        SHand As BodyPartVars
'/E\---------------------------------Mob body parts----------------------------------
        
        Bleeding As Integer 'The amount of blood mob bleeds every delayed.Bleeding
        Delay As DelayVars
        
        Items As String
        Wear As String
        Damage As Integer 'The damage mob does on a hit
        locX As Integer ' \
        locY As Integer '  > Location on the current area (X,Y,Z)
        locZ As Integer ' /
        ApproachedPCs As String 'PCs that approached the mob
        ApproachedMobs As String 'Mobs that approached the mob
        Area As String 'Area name
        AI As AIVars
    End Type
    
    Public PrototypeMob() As MobVars 'Keeps and stores the prototype of each mob type (ID)
    Public Mob() As MobVars 'Stores the ACTIVE mobs (Removes them when they are killed)
    Public MFreeVNums As String 'Has the VNums of mobs that got free (e.g If mob died)
'< Mobile vars

Function OpenMobEmoteTags(ByVal Str$, Index%, Arguement$, Mnum%) As String
'\B/------------------------------Open Mob related tags------------------------------
     Str = Replace(Str, "<target>", Mob(Mnum).Name): Arguement = Replace(LCase(Arguement), LCase(Mob(Mnum).Name), "")
    Str = Trim(Replace(Str, "<arg>", Arguement))
    If InStr(Str, "<hisher>") Then Str = Replace(Str, "<hisher>", HeShe(Char(Index).Gender, "HisHer"))
    OpenMobEmoteTags = Str & "."
'/E\------------------------------Open Mob related tags------------------------------
End Function
Sub MobDeath(Mnum%)
    'Transmiting the death message to surrounding players
    If IsApproachedMob(Mnum) Then Call ApproachRemovalMob(Mnum)
    Call TransmitLocalMob(Mnum, Mob(Mnum%).Name & " was slain. Think about it. You just killed a living creature.")
    Call UnloadMob(Mnum)
End Sub
Sub TransmitLocalMob(Mnum%, ByVal Transmit As String, Optional NoSend As Integer = -1, Optional Color As String = WHITE)
    Dim I%, WrappedText$
    Dim Arr As Variant
    Arr = StringToArray(PCsNearMob(Mnum%))
    If UBound(Arr) >= 0 Then
        WrappedText = Wrapped(Transmit, Color)
        For I = LBound(Arr) To UBound(Arr)
            If Arr(I) <> NoSend Then Send Val(Arr(I)), WrappedText, , False
        Next I
    End If
    End Sub
Function PCsNearMob(Mnum%) As String
    PCsNearMob = Area(Mob(Mnum%).Area) _
      .Room(Mob(Mnum%).locX, Mob(Mnum%).locY, Mob(Mnum%).locZ).PCs
End Function
Function MobIsHere(Index As Integer, ByVal Name As String) As Integer
'Returns the VNum if mob found
    Dim I%, Lastword$
    Dim Arr As Variant
    
    Arr = StringToArray(Mobs(Index))
        For I = LBound(Arr) To UBound(Arr)
            If InStr(QteMe(LCase(Mob(Arr(I)).Aliases)), QteMe(Name)) > 0 Then
                    MobIsHere = Val(Arr(I))
                    Exit Function
            End If
        Next I
End Function
Sub GetMobLook(ByVal Vnum As Variant, ByRef Condition$, ByRef WearedItemsLook$)

With Mob(Vnum)
    '\B/----------------------------------------See wearing equipment----------------------------------------
       WearedItemsLook = WearsItems(SortItemsList(.Wear), _
                                    .PHand.Name, .Gender)
    '/E\----------------------------------------See wearing equipment----------------------------------------
    '\B/--------------------------------------------See condition--------------------------------------------
        Dim HisHer$
        HisHer = HeShe(.Gender, "hisher")
        If .Head.Cond < .MaxHP Then _
          Condition = Condition & HisHer & " head is " & GetCondition(.Head.Cond, .MaxHP) & ", "
        If .Torso.Cond < .MaxHP Then _
          Condition = Condition & HisHer & " torso is " & GetCondition(.Torso.Cond, .MaxHP) & ", "
        If .PHand.Cond < .MaxHP Then _
          Condition = Condition & HisHer & " " & .PHand.Name & " is " & GetCondition(.PHand.Cond, .MaxHP) & ", "
        If .SHand.Cond < .MaxHP Then _
          Condition = Condition & HisHer & " " & .SHand.Name & " is " & GetCondition(.SHand.Cond, .MaxHP) & ", "
        If .Legs.Cond < .MaxHP Then _
          Condition = Condition & HisHer & " legs are " & GetCondition(.Legs.Cond, .MaxHP) & ", "
        If Condition <> "" Then Condition = Proper(Left(Condition, Len(Condition) - 2)) & "."
End With
    '/E\--------------------------------------------See condition--------------------------------------------
End Sub
Function ActivateMob(ID As Integer) As Integer
'This function creates the mob with the mob in the Mob() array and returns its vnum
    Dim FreeVnum As Integer
    If Len(MFreeVNums) > 0 Then
        FreeVnum = GetFreeVNum(MFreeVNums)
    ElseIf Mob(UBound(Mob)).ID > 0 Then
        FreeVnum = UBound(Mob) + 1
        ReDim Preserve Mob(1 To FreeVnum)
    Else
        FreeVnum = 1
    End If
    
    Load frmMain.tmrBleedingMob(FreeVnum)
    Load frmMain.tmrDelayedComsMob(FreeVnum)
        
    If ID > UBound(PrototypeMob) Then ActivateMob = 0: Exit Function 'Error
    Mob(FreeVnum) = PrototypeMob(ID)
    ActivateMob = FreeVnum
End Function
Function LoadMob(X%, Y%, Z%, Area%, Index As Integer, ID As Integer) As Integer
    Dim Vnum As Integer
    'Adds the mob to Index's location
    Vnum = ActivateMob(ID)
    If Vnum > 0 Then Call AddMob(Index, Vnum) Else LoadMob = 0: Exit Function 'Error
    
'\B/----------------Setting the location of the mob in Mob's records----------------
    Mob(Vnum).locX = X
    Mob(Vnum).locY = Y
    Mob(Vnum).locZ = Z
    Mob(Vnum).Area = Area
'/E\----------------Setting the location of the mob in Mob's records----------------
    
    LoadMob = Vnum
End Function
Function UnloadMob(Vnum%)

    Unload frmMain.tmrBleedingMob(Vnum)
    Unload frmMain.tmrDelayedComsMob(FreeVnum)
    
    Call RemoveMob(Vnum%)
    MFreeVNums = AddToString(MFreeVNums, Vnum)
End Function
Sub ApproachRemovalMob(Vnum As Integer)
With Mob(Vnum)

Dim Arr As Variant, I As Integer

'\B/------------------------If approached PCs then retreating------------------------
    If Len(.ApproachedPCs) > 0 Then
        Arr = StringToArray(.ApproachedPCs)
        For I = LBound(Arr) To UBound(Arr)
            Char(Arr(I)).ApproachedMobs = RemoveFromString(Char(Arr(I)).ApproachedPCs, Vnum)
        Next I
        .ApproachedPCs = ""
    End If
'/E\------------------------If approached PCs then retreating------------------------

'\B/-----------------------If approached mobs then retreating-----------------------
    If Len(.ApproachedMobs) > 0 Then
        Arr = StringToArray(.ApproachedMobs)
        For I = LBound(Arr) To UBound(Arr)
            Mob(Arr(I)).ApproachedPCs = _
              RemoveFromString(Mob(Arr(I)).ApproachedPCs, Vnum)
        Next I
        .ApproachedMobs = ""
    End If
'/E\-----------------------If approached mobs then retreating-----------------------

End With
End Sub
Function IsApproachedMob(Vnum%) As Boolean
With Mob(Vnum)

    If Len(.ApproachedPCs) + Len(.ApproachedMobs) > 0 Then _
      IsApproachedMob = True Else IsApproachedMob = False

End With
End Function
Sub AttackMob(Mnum%, PreparingTarget$, PreparingOthers$, _
                     DelayedCommand$, Optional CheckForApproach As Boolean)
    Dim NameIndex%, TargetMnum%
    
    With Mob(Mnum)
    NameIndex = .Delay.PCTarget
    TargetMnum = .Delay.MobItemVnum
    
'\B/----------------------------------Here at all?----------------------------------
    If Not CheckMPI(NameIndex) And Not CheckMPI(TargetMnum) Then
        Call RemoveDelayMob(Mnum)
        Exit Sub
    End If
'/E\----------------------------------Here at all?----------------------------------
        
    If CheckMPI(NameIndex) Then
        
        With Char(NameIndex)
        PreparingTarget = GlobalOpenTags(PreparingTarget, Index:=NameIndex)
        PreparingOthers = GlobalOpenTags(PreparingOthers, Index:=NameIndex)
        End With
        
'\B/-----------------------------------Approached?-----------------------------------
        If InStr(QteMe(Mob(Mnum).ApproachedPCs), QteMe(NameIndex)) = 0 Then Exit Sub
'/E\-----------------------------------Approached?-----------------------------------
        Send NameIndex, PreparingTarget
        TransmitLocalMob Mnum, PreparingOthers, NameIndex
    
    ElseIf CheckMPI(Mnum) Then
        
        With Mob(TargetMnum)
        PreparingOthers = GlobalOpenTags(PreparingOthers, Mnum:=Mnum)
        End With
        
'\B/-----------------------------------Approached?-----------------------------------
        If InStr(QteMe(Mob(Mnum).ApproachedMobs), QteMe(TargetMnum)) = 0 Then Exit Sub
'/E\-----------------------------------Approached?-----------------------------------
        TransmitLocalMob Mnum, PreparingOthers
    End If

    Call AddDelayMob(Mnum, Seconds:=2, CommandName:=DelayedCommand, TargetMnum:=TargetMnum, NameIndex:=NameIndex)
    End With

End Sub
Sub DelayedAttackMob(Mnum%, HittedBodyPart As BodyPartVars, _
             Optional MissMsgOthers$ = "", Optional MissMsgTarget$, _
             Optional MsgOthers$ = "", Optional MsgTarget$ = "")
    Dim DamageDone%
    Dim NameIndex% 'Holds the found Index of a Name
    Dim TargetMnum As Integer 'Holds the found Vnum of the mob
    
    NameIndex = Mob(Mnum).Delay.PCTarget
    TargetMnum = Mob(Mnum).Delay.MobItemVnum
        
'\B/-------------------------------Is the target here?-------------------------------
    If TargetChangedMob(Mnum, NameIndex, TargetMnum) Then
        Call RemoveDelayMob(Mnum)
        Exit Sub
    End If
'/E\-------------------------------Is the target here?-------------------------------

'\B/-----------------------------------PC to PC hit-------------------------------------
    If CheckMPI(NameIndex) Then
        
        With Char(NameIndex)
            MissMsgTarget = GlobalOpenTags(MissMsgTarget, Index:=NameIndex, BodyPartName:=HittedBodyPart.Name)
            MissMsgOthers = GlobalOpenTags(MissMsgOthers, Index:=NameIndex, BodyPartName:=HittedBodyPart.Name)
            
            MsgTarget = GlobalOpenTags(MsgTarget, Index:=NameIndex, BodyPartName:=HittedBodyPart.Name)
            MsgOthers = GlobalOpenTags(MsgOthers, Index:=NameIndex, BodyPartName:=HittedBodyPart.Name)
        End With
        
        If ACDefendedMob(Mnum, NameIndex, Mob(Mnum).Damage, HittedBodyPart, _
           MissMsgTarget, MissMsgOthers) Then _
           Call AIMissTargetMob(Mnum, TargetMnum:=TargetMnum): Exit Sub
        
        DamageDone = DoDamage(NameIndex, Mob(Mnum).Damage, HittedBodyPart)
        Call AIHitTargetMob(Mnum, DamageDone, NameIndex)
        
        
    '\B/-----------------------------------Transmiting-----------------------------------
        Send NameIndex, MsgTarget
        TransmitLocalMob Mnum, MsgOthers, NameIndex
    '/E\-----------------------------------Transmiting-----------------------------------
    End If
'/E\-----------------------------------PC TO PC hit-------------------------------------
    
    '\B/-----------------------------------PC to MOB hit-------------------------------------
    If CheckMPI(TargetMnum) Then
        
        With Mob(TargetMnum)
        MissMsgTarget = GlobalOpenTags(MissMsgTarget, Mnum:=TargetMnum, BodyPartName:=HittedBodyPart.Name)
        MissMsgOthers = GlobalOpenTags(MissMsgOthers, Mnum:=TargetMnum, BodyPartName:=HittedBodyPart.Name)
        
        MsgTarget = GlobalOpenTags(MsgTarget, Mnum:=TargetMnum, BodyPartName:=HittedBodyPart.Name)
        MsgOthers = GlobalOpenTags(MsgOthers, Mnum:=TargetMnum, BodyPartName:=HittedBodyPart.Name)
        End With
        
        If ACDefendedMob(Mnum, NameIndex, Mob(Mnum).Damage, HittedBodyPart, _
          MissMsgTarget, MissMsgOthers) Then _
            Call AIMissTargetMob(Mnum, TargetMnum:=TargetMnum): Exit Sub
           
        DamageDone = DoDamageMob(TargetMnum, Mob(Mnum).Damage, HittedBodyPart)
        Call AIHitTargetMob(Mnum, DamageDone, TargetMnum:=TargetMnum)
        
        TransmitLocalMob Mnum, MsgOthers

    End If
    '/E\-----------------------------------PC TO MOB hit-------------------------------------
    
    'Updating the bodypart
    With Mob(Mnum).Delay
        Call LetBodyPart(.PCTarget, .MobItemVnum, HittedBodyPart)
    End With
End Sub
Function ACDefendedMob(Mnum%, NameIndex%, Damage As Integer, HittedBodyPart As BodyPartVars, _
                    Optional MissMsgTarget$ = "", Optional MissMsgOthers$ = "") As Boolean
    Damage = 1 'TEMP
    ACDefendedMob = False
    If Damage <= HittedBodyPart.AC Then
        ACDefendedMob = True
        If NameIndex > 0 Then Send NameIndex, MissMsgTarget
        TransmitLocalMob Mnum, MissMsgOthers, NameIndex
    End If
End Function
