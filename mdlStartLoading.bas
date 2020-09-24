Attribute VB_Name = "mdlStartLoading"
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
Sub LoadFromDatabaseToMemory()
    
'\B/-----------------------------Settings ready records-----------------------------
    Set CharRecord.SnapShot = DB.OpenRecordset("Characters", dbOpenSnapshot)  'Making the characters' RecordSet
    Set CharRecord.Dynaset = DB.OpenRecordset("Characters", dbOpenDynaset)   'Making the characters' RecordSet
    CharRecord.Table = "Characters"
'/E\-----------------------------Settings ready records-----------------------------
    
'\B/----------------------------Adding the player status----------------------------
    GDic.Add "admin", Admin
    GDic.Add "immortal", Immortal
    GDic.Add "mortal", Mortal
'/E\----------------------------Adding the player status----------------------------
    
    Call LoadItemsToMemory
    
    Call LoadMobsToMemory
    
    Call LoadAreasToMemory
    
    Call LoadEmotesToMemory

End Sub


Sub LoadItemsToMemory()

 Dim ID%
 
'\B/-----------------------------Loading all the item prototypes to memory--------------------
On Error Resume Next
    Set GRecord = DB.OpenRecordset("Items", dbOpenSnapshot)
    If GRecord.EOF = False Then GRecord.MoveLast
    ReDim Preserve PrototypeItem(1 To GRecord.RecordCount)
    If GRecord.BOF = False Then GRecord.MoveFirst
        Do
            ID = GRecord!ID
            With PrototypeItem(ID)
            
            .ID = ID
            .Name = GRecord!Name
            'Adding item names to forbidden names
            If InStr(QteMe(ForbiddenNames), QteMe(GRecord!Name)) = 0 Then _
              ForbiddenNames = AddToString(ForbiddenNames, GRecord!Name)
            .Description = GRecord!Description
            .Type = GRecord!Type
            .Subtype = GRecord!Subtype
            .Wear = GRecord!Wear
            .Damage = GRecord!Damage
            .AC = GRecord!AC
            'Adding aliases
            .Aliases = AddToString(.Aliases, _
              .Name & QteMe(.Type) & .Subtype)
            GRecord.MoveNext
            
            End With
        Loop Until GRecord.EOF
'/E\------------------------------Loading all the item prototypes to memory------------------------------------------
        
'\B/-------------------------Creating bare hands as VNum 1!-------------------------
        ReDim Item(1 To 1)
        Call ActivateItem(4)
'/E\-------------------------Creating bare hands as VNum 1!-------------------------
    
End Sub
Sub LoadMobsToMemory()
'\B/-----------------------------Loading all the mob prototypes to memory------------------------------------------
    Dim TempWear$, TempItems$
    Dim ID%, Correction%, Vnum%
    
    Set GRecord = DB.OpenRecordset("Mobs", dbOpenSnapshot)
    If GRecord.EOF = False Then GRecord.MoveLast
    ReDim Preserve PrototypeMob(1 To GRecord.RecordCount)
    If GRecord.BOF = False Then GRecord.MoveFirst
    Do
        ID = GRecord!ID
        PrototypeMob(ID).ID = ID
        PrototypeMob(ID).Name = GRecord!Name
        PrototypeMob(ID).Description = GRecord!Description
        PrototypeMob(ID).Gender = GRecord!Gender
        PrototypeMob(ID).MaxHP = GRecord!HPMax
        PrototypeMob(ID).HP = PrototypeMob(ID).MaxHP
        
        Dim Cond$, AC$
        
        Cond = GRecord!BodyPartsCondition
        AC = GRecord!BodyPartsAC
        
        PrototypeMob(ID).Head.Cond = GetWordByNum(1, Cond)
        PrototypeMob(ID).Torso.Cond = GetWordByNum(2, Cond)
        PrototypeMob(ID).Legs.Cond = GetWordByNum(3, Cond)
        PrototypeMob(ID).PHand.Cond = GetWordByNum(4, Cond)
        PrototypeMob(ID).SHand.Cond = GetWordByNum(5, Cond)
        
        PrototypeMob(ID).Head.AC = GetWordByNum(1, AC)
        PrototypeMob(ID).Torso.AC = GetWordByNum(2, AC)
        PrototypeMob(ID).Legs.AC = GetWordByNum(3, AC)
        PrototypeMob(ID).PHand.AC = GetWordByNum(4, AC)
        PrototypeMob(ID).SHand.AC = GetWordByNum(5, AC)
        
        PrototypeMob(ID).Head.Name = "head"
        PrototypeMob(ID).Torso.Name = "torso"
        PrototypeMob(ID).Legs.Name = "legs"
        PrototypeMob(ID).PHand.Name = "right hand"
        PrototypeMob(ID).SHand.Name = "left hand"

'        PrototypeMob(Id).Type = GRecord!dbMOBtype)
'        PrototypeMob(Id).Subtype = GRecord!dbMOBsubtype)
        TempWear = GRecord!Wear & ""
        TempItems = GRecord!Items & ""
   
        For Correction = LBound(StringToArray(TempWear)) To UBound(StringToArray(TempWear))
            Vnum = ActivateItem(Val(GetWordByNum(2, (Split(TempWear, ",")(Correction)), " ")))
            PrototypeMob(ID).Wear = AddToString(PrototypeMob(ID).Wear, Item(Vnum).Wear & " " & Vnum)
            Call WearItemMob(PrototypeMob(ID), Vnum)
        Next Correction
            If PrototypeMob(ID).PHand.WearVnum = 0 Then PrototypeMob(ID).PHand.WearVnum = 1
        
        For Correction = LBound(StringToArray(TempItems)) To UBound(StringToArray(TempItems))
            Vnum = ActivateItem(Val(Split(TempItems, ",")(Correction)))
            PrototypeMob(ID).Items = AddToString(PrototypeMob(ID).Items, Vnum)
        Next Correction
        
        'Adding aliases
        PrototypeMob(ID).Aliases = AddToString(PrototypeMob(ID).Aliases, PrototypeMob(ID).Name)

        'Adding Mob names to forbidden names
        If InStr(QteMe(ForbiddenNames), QteMe(GRecord!Name)) = 0 Then _
          ForbiddenNames = AddToString(ForbiddenNames, GRecord!Name)
        
        GRecord.MoveNext
    Loop Until GRecord.EOF
'/E\-----------------------------Loading all the mob prototypes---------------------------------
End Sub
Sub LoadAreasToMemory()

    Dim AreaName As String
    Dim Vnum As String

'\B/---------------------------------Gets the areas list-----------------------------
    Dim HighestX%, HighestY%, HighestZ%
    Dim LowestX%, LowestY%, LowestZ%
    Dim I%, ID%, Correction%
     
     Open App.Path & "\world.dat" For Input As #1
        Do
            Inc I
            Input #1, AreaName
            AreaName = "Map_" & AreaName
            ReDim Preserve Area(I)
            Area(I).Name = AreaName
    '/E\----------------------------------Gets the areas list----------------------------
    '\B/------------------------Getting the top locations for the ReDim------------------------------------------
            Set GRecord = DB.OpenRecordset("SELECT MAX(X) FROM " & AreaName, dbOpenSnapshot)
            HighestX = GRecord(0)
            Set GRecord = DB.OpenRecordset("SELECT MIN(X) FROM " & AreaName, dbOpenSnapshot)
            LowestX = GRecord(0)
            Set GRecord = DB.OpenRecordset("SELECT MAX(Y) FROM " & AreaName, dbOpenSnapshot)
            HighestY = GRecord(0)
            Set GRecord = DB.OpenRecordset("SELECT MIN(Y) FROM " & AreaName, dbOpenSnapshot)
            LowestY = GRecord(0)
            Set GRecord = DB.OpenRecordset("SELECT MAX(Z) FROM " & AreaName, dbOpenSnapshot)
            HighestZ = GRecord(0)
            Set GRecord = DB.OpenRecordset("SELECT MIN(Z) FROM " & AreaName, dbOpenSnapshot)
            LowestZ = GRecord(0)
            ReDim Area(I).Room(LowestX To HighestX, LowestY To HighestY, LowestZ To HighestZ)
    '/E\-------------------------Getting the top locations for the ReDim------------------------------------------
    '\B/-------------------------Loading all the areas in list to memory------------------------------------------
            Set GRecord = DB.OpenRecordset(AreaName, dbOpenSnapshot)
            GRecord.MoveFirst
            Do
            With Area(I).Room(GRecord!X, GRecord!Y, GRecord!Z)
            
                Dim TempMobs$, TempItems$
            
                .Description = GRecord!Description & ""
                .Exits = GRecord!Exits & ""
                TempMobs = GRecord!Mobs & ""
                TempItems = GRecord!Items & ""
        
                For Correction = LBound(StringToArray(TempMobs)) To UBound(StringToArray(TempMobs))
                    ReDim Mob(1 To 1)
                    Vnum = ActivateMob(Val(Split(TempMobs, ",")(Correction)))
                    .Mobs = AddToString(.Mobs, Vnum)
                    Mob(Vnum).locX = GRecord!X
                    Mob(Vnum).locY = GRecord!Y
                    Mob(Vnum).locZ = GRecord!Z
                    Mob(Vnum).Area = I
                Next Correction
                
                For Correction = LBound(StringToArray(TempItems)) To UBound(StringToArray(TempItems))
                    Vnum = ActivateItem(Val(Split(TempItems, ",")(Correction)))
                    .Items = AddToString(.Items, Vnum)
                Next Correction
                
                GRecord.MoveNext
            
            End With
            Loop Until GRecord.EOF
        Loop Until EOF(1)
     Close #1
'/E\-------------------------Loading all the areas in list to memory------------------------------------------

End Sub
Sub LoadEmotesToMemory()
 
    Dim I%, ID%

'\B/----------------------------Loading all the emotes to memory--------------------------------------------
    Set GRecord = DB.OpenRecordset("Emotes", dbOpenSnapshot)
    If GRecord.EOF = False Then GRecord.MoveLast
    ReDim Emotes(1 To GRecord.RecordCount)
    If GRecord.BOF = False Then GRecord.MoveFirst
        Do
            Inc I
            Emotes(I).ID = GRecord!ID
            ForbiddenNames = AddToString(ForbiddenNames, GRecord!ID)
            Emotes(I).Self = GRecord!Self
            Emotes(I).Others = GRecord!Others
            Emotes(I).SelfTarget = GRecord!SelfTarget
            Emotes(I).OthersTarget = GRecord!OthersTarget
            Emotes(I).Target = GRecord!Target
            GRecord.MoveNext
        Loop Until GRecord.EOF
'/E\---------------------------Loading all the emotes to memory------------------------------------------
    
    End Sub

