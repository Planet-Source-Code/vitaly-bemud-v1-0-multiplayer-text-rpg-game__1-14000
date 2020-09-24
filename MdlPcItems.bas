Attribute VB_Name = "MdlPCItems"
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
Sub WearItem(Index%, Vnum%)
Dim ReplaceString$
With Char(Index)
    
    Select Case Item(Vnum).Wear
    Case "torso"
        .Torso.AC = .Torso.AC + Item(Vnum).AC
        .Torso.WearVnum = Vnum
    Case "legs"
        .Legs.AC = .Legs.AC + Item(Vnum).AC
        .Legs.WearVnum = Vnum
    Case "head"
        .Head.AC = .Head.AC + Item(Vnum).AC
        .Head.WearVnum = Vnum
    Case "phand"
        .PHand.AC = .PHand.AC + Item(Vnum).AC
        .PHand.WearVnum = Vnum
        .Damage = .Damage + Item(Vnum).Damage
        '..PHand.WearVnum = Vnum%: ..PHand.WearIndex = ItemIndex%
    Case "shand"
        .SHand.AC = .SHand.AC + Item(Vnum).AC
        .SHand.WearVnum = Vnum
        .Damage = .Damage + Item(Vnum).Damage
    End Select
    
    'This time the ReplaceString value is in a temp var to make the code more clean (AHEM)
    ReplaceString = Item(Vnum).Wear & " " & Vnum
   
    If Len(.Wear) > 0 Then _
      .Wear = ReplaceString & "," & .Wear _
      Else: .Wear = ReplaceString

End With
End Sub
Sub RemoveWearItem(Index%, Vnum%)
With Char(Index)
    Dim ReplaceString$
    Select Case Item(Vnum).Wear
    Case "torso"
        .Torso.AC = .Torso.AC - Item(Vnum).AC
        .Torso.WearVnum = 0
    Case "legs"
        .Legs.AC = .Legs.AC - Item(Vnum).AC
        .Legs.WearVnum = 0
    Case "head"
        .Head.AC = .Head.AC - Item(Vnum).AC
        .Head.WearVnum = 0
    Case "phand"
        .PHand.AC = .PHand.AC - Item(Vnum).AC
        .PHand.WearVnum = 1
        .Damage = .Damage - Item(Vnum).Damage
    Case "shand"
        .SHand.AC = .SHand.AC - Item(Vnum).AC
        .SHand.WearVnum = 0
        .Damage = .Damage - Item(Vnum).Damage
    End Select
    
    'This time the ReplaceString value is in a temp var to make the code more clean (AHEM)
    ReplaceString = Item(Vnum).Wear & " " & Vnum
    
    .Wear = RemoveFromString(.Wear, ReplaceString)

End With
End Sub
Function ItemIsWeared(Index As Integer, Name As String) As Integer
With Char(Index)

'Returns an ARRAY with 1=ItemID, 2=ItemIndex, 3=Where is it weared
    Dim I%, Vnum%
    Dim Arr As Variant, WhichOne%, Lastword$
    Lastword = GetWordByNum(CountWords(Name), Name, " ")
'\B/----------------------------------WhichOne code----------------------------------
    If IsNumeric(Lastword) Then WhichOne = Val(Lastword): _
      Name = Replace(Name, " " & WhichOne, "") Else WhichOne = 1
'/E\----------------------------------WhichOne code----------------------------------
    
    Arr = StringToArray(.Wear)
        For I = LBound(Arr) To UBound(Arr)
            Vnum = GetWordByNum(2, Arr(I), " ")

            If InStr(QteMe(LCase(Item(Vnum).Aliases)), QteMe(Name)) > 0 Then
                WhichOne = WhichOne - 1
                If WhichOne = 0 Then
                    ItemIsWeared = Vnum
                    Exit Function
                End If
            End If
        
        Next I

End With
End Function
Sub DropAll(Index%)
With Char(Index)
    
    'Drops all the equipment on the ground
    Area(.Area).Room(.locX, .locY, .locZ).Items = _
      AddToString(Area(.Area).Room(.locX, .locY, .locZ).Items, .Items)

End With
End Sub
Sub RemoveAll(Index%)
    'Removes all the weared equipment
    Dim I%, Arr$()
    Arr = Split(Char(Index).Wear, ",")
    For I = LBound(Arr) + 1 To UBound(Arr) Step 2
        Char(Index).Items = AddToString(Char(Index).Items, Arr(I))
    Next I
End Sub
Function ItemIsHere(Index As Integer, ByVal Name As String) As Integer
'Returns the vnum of the item
    Dim I%, Lastword$
    Dim Arr As Variant, WhichOne%
    
    Lastword = GetWordByNum(CountWords(Name), Name, " ")
    If IsNumeric(Lastword) Then WhichOne = Val(Lastword): _
      Name = Replace(Name, " " & WhichOne, "") Else WhichOne = 1
    Arr = StringToArray(Items(Index))
        For I = LBound(Arr) To UBound(Arr)
            If InStr(QteMe(LCase(Item(Arr(I)).Aliases)), QteMe(Name)) > 0 Then
                WhichOne = WhichOne - 1
                If WhichOne = 0 Then
                    ItemIsHere = Arr(I)
                    Exit Function
                End If
            End If
        Next I
End Function
Sub AddItemInv(Index%, Vnum%)
Dim ReplaceString$
    
    If Len(Char(Index).Items) > 0 Then _
      Char(Index).Items = Vnum & "," & Char(Index).Items _
      Else: Char(Index).Items = Vnum
End Sub
Sub RemoveItemInv(Index%, Vnum%)
Dim ReplaceString$
    
    If Trim(Vnum) = Char(Index).Items Then Char(Index).Items = "": Exit Sub
    Char(Index).Items = Replace(QteMe(Char(Index).Items), QteMe(Vnum), ",")
    Char(Index).Items = Mid(Char(Index).Items, 2, Len(Char(Index).Items) - 2)
End Sub
Function ItemIsInInv(Index As Integer, Name As String) As Integer
'Returns an ARRAY with 1=ItemID, 2=ItemIndex
    Dim I%, Vnum%
    Dim Arr As Variant, WhichOne%, Lastword$
    Lastword = GetWordByNum(CountWords(Name), Name, " ")
    
'\B/--------------Checks if it is a request on not the first found item--------------
    If IsNumeric(Lastword) Then WhichOne = Val(Lastword): _
      Name = Replace(Name, " " & WhichOne, "") Else WhichOne = 1
'/E\--------------Checks if it is a request on not the first found item--------------
    
    Arr = StringToArray(Char(Index).Items)
        For I = LBound(Arr) To UBound(Arr)
            If InStr(QteMe(LCase(Item(Arr(I)).Aliases)), QteMe(Name)) > 0 Then
                
                WhichOne = WhichOne - 1
                
                If WhichOne = 0 Then
                    ItemIsInInv = Arr(I)
                    Exit Function
                End If
            End If
        Next I
End Function


