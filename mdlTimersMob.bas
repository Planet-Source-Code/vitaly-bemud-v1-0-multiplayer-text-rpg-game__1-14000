Attribute VB_Name = "mdlMobTimers"
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
Sub DelayedComsMob(Mnum As Integer)
With Mob(Mnum)
    Select Case .Delay.Command
'******************************************* A *****************************************
'    Case "approach-d"
'        Call mcmdApproachDelay(Mnum)
'******************************************* B *****************************************
'******************************************* C *****************************************
'******************************************* D *****************************************
'******************************************* E *****************************************
'******************************************* F *****************************************
'******************************************* G *****************************************
'******************************************* H *****************************************
    Case "hit"
        Call mcmdHit(Mnum)
    Case "hit-d"
        Call mcmdHitDelay(Mnum)
'******************************************* I *****************************************
'******************************************* J *****************************************
'******************************************* K *****************************************
'******************************************* K *****************************************
'******************************************* L *****************************************
'******************************************* M *****************************************
'******************************************* N *****************************************
'******************************************* O *****************************************
'******************************************* P *****************************************
'******************************************* Q *****************************************
'******************************************* R *****************************************
'    Case "retreat-d"
'        Call mcmdRetreatDelay(Mnum)
'******************************************* S *****************************************
'******************************************* T *****************************************
'******************************************* U *****************************************
'******************************************* V *****************************************
'******************************************* W *****************************************
'******************************************* X *****************************************
'******************************************* Y *****************************************
'******************************************* Z *****************************************
    End Select
End With
End Sub
Sub BleedingTimerMob(Vnum As Integer)
    If Mob(Vnum).HP - Mob(Vnum).Bleeding > 0 Then
        Mob(Vnum).HP = Mob(Vnum).HP - Mob(Vnum).Bleeding
        TransmitLocalMob Vnum, Mob(Vnum).Name & bRED & " bleeds" & WHITE & "!"
    Else
        Call MobDeath(Vnum) 'Mob dies
    End If
End Sub
Sub AddDelayMob(Mnum As Integer, Seconds As Integer, CommandName As String, _
            Optional TargetMnum = 0, Optional NameIndex As Integer = 0)
With Mob(Mnum)
    
    .Delay.Busy = True
    .Delay.Command = CommandName
    .Delay.PCTarget = NameIndex
    .Delay.MobItemVnum = TargetMnum
    
    frmMain.tmrDelayedComsMob(Mnum).Enabled = True
    frmMain.tmrDelayedComsMob(Mnum).Interval = Seconds * IntervalFormat

End With
End Sub
Sub RemoveDelayMob(Mnum As Integer)
With Mob(Mnum)
    
    frmMain.tmrDelayedComsMob(Mnum).Enabled = False
    .Delay.PCTarget = 0
    .Delay.MobItemVnum = Empty
    .Delay.Busy = False

End With
End Sub

Function TargetChangedMob(Mnum%, NameIndex%, TargetMnum%) As Boolean
Dim TargetNotHere As Boolean

TargetChangedMob = False

'\B/------------------------Checking the target is still here------------------------
        If CheckMPI(TargetMnum) And InStr(QteMe(Mob(Mnum).ApproachedMobs), QteMe(TargetMnum)) = 0 Then TargetNotHere = True
        If CheckMPI(NameIndex) And InStr(QteMe(Mob(Mnum).ApproachedPCs), QteMe(NameIndex)) = 0 Then TargetNotHere = True
'/E\------------------------Checking the target is still here------------------------
            
'\B/---------------------------------Target is gone---------------------------------
        If TargetNotHere Then
            RemoveDelayMob Mnum
            TargetChangedMob = True
        End If
'/E\---------------------------------Target is gone---------------------------------
End Function

