Attribute VB_Name = "mdlMobCommands"
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
Sub mcmdHit(Mnum%)
Dim DelayCommand$
Dim PreparingSelf$, PreparingTarget$, PreparingOthers$

With Mob(Mnum)
PreparingSelf = "You aim to hit <targetname>."
PreparingTarget = .Name & " aims to hit you!"
PreparingOthers = .Name & " aims to hit <targetname>."
End With

Call AttackMob(Mnum, PreparingTarget, PreparingOthers, "hit-d")

End Sub
Sub mcmdHitDelay(Mnum%)
    Dim MissMsgOthers$, MissMsgTarget$
    Dim MsgOthers$, MsgTarget$
    Dim HitBodyPart As BodyPartVars
    
    With Mob(Mnum).Delay
        HitBodyPart = GetBodyPart(.PCTarget, .MobItemVnum)
    End With
    
    With Mob(Mnum)
    MissMsgTarget = .Name & " misses you."
    MissMsgOthers = .Name & " misses <targetname>."
    
    MsgTarget = .Name & " hits you on <bodypartname>!"
    MsgOthers = .Name & " hits <targetname> on <bodypartname>!"
    End With
    
    Call DelayedAttackMob(Mnum, HitBodyPart, _
                MissMsgOthers, MissMsgTarget, _
                MsgOthers, MsgTarget)
End Sub

