Attribute VB_Name = "mdlMobItems"
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
Sub WearItemMob(ByRef Mob As MobVars, Inum%)
Dim ReplaceString$
With Mob

    Select Case Item(Inum).Wear
    Case "torso"
        .Torso.AC = .Torso.AC + Item(Inum).AC
        .Torso.WearVnum = Inum
    Case "legs"
        .Legs.AC = .Legs.AC + Item(Inum).AC
        .Legs.WearVnum = Inum
    Case "head"
        .Head.AC = .Head.AC + Item(Inum).AC
        .Head.WearVnum = Inum
    Case "phand"
        .PHand.AC = .PHand.AC + Item(Inum).AC
        .PHand.WearVnum = Inum
        .Damage = .Damage + Item(Inum).Damage
        '..PHand.WearVnum = Inum%: ..PHand.WearIndex = ItemIndex%
    Case "shand"
        .SHand.AC = .SHand.AC + Item(Inum).AC
        .SHand.WearVnum = Inum
        .Damage = .Damage + Item(Inum).Damage
    End Select
    
    'This time the ReplaceString value is in a temp var to make the code more clean (AHEM)
    ReplaceString = Item(Inum).Wear & " " & Inum
   
    If Len(.Wear) > 0 Then _
      .Wear = ReplaceString & "," & .Wear _
      Else: .Wear = ReplaceString

End With
End Sub

