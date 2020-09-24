Attribute VB_Name = "mdlMobAI"
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
Sub AIGotHitMob(Mnum%, DamageDone%, Optional NameIndex% = 0, Optional TargetMnum% = 0)
With Mob(Mnum)
    
    Call AddDelayMob(Mnum, 2, "hit", TargetMnum, NameIndex)

End With
End Sub
Sub AIMissTargetMob(Mnum%, Optional NameIndex% = 0, Optional TargetMnum% = 0)
With Mob(Mnum)
    
    Call AddDelayMob(Mnum, 2, "hit", TargetMnum, NameIndex)

End With
End Sub
Sub AIHitTargetMob(Mnum%, DamageDone%, Optional NameIndex% = 0, Optional TargetMnum% = 0)
With Mob(Mnum)
    
    Call AddDelayMob(Mnum, 2, "hit", TargetMnum, NameIndex)

End With
End Sub

