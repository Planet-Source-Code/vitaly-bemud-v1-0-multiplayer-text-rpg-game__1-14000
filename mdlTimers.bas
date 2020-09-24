Attribute VB_Name = "mdlPcTimers"
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
Sub DelayedCommands(Index As Integer)
    Select Case Char(Index).Delay.Command
'******************************************* A *****************************************
    Case "approach-d"
        Call pcmdApproachDelay(Index)
'******************************************* B *****************************************
'******************************************* C *****************************************
'******************************************* D *****************************************
'******************************************* E *****************************************
'******************************************* F *****************************************
'******************************************* G *****************************************
'******************************************* H *****************************************
    Case "hit-d"
        Call pcmdHitDelay(Index)
'******************************************* I *****************************************
'******************************************* J *****************************************
'******************************************* K *****************************************
'******************************************* K *****************************************
'******************************************* L *****************************************
'******************************************* M *****************************************
    Case "missile-d"
        Call pcmdMissileDelay(Index)
'******************************************* N *****************************************
'******************************************* O *****************************************
'******************************************* P *****************************************
'******************************************* Q *****************************************
'******************************************* R *****************************************
    Case "retreat-d"
        Call pcmdRetreatDelay(Index)
'******************************************* S *****************************************
'******************************************* T *****************************************
'******************************************* U *****************************************
'******************************************* V *****************************************
'******************************************* W *****************************************
'******************************************* X *****************************************
'******************************************* Y *****************************************
'******************************************* Z *****************************************
    End Select
    Call RemoveDelay(Index)
End Sub
Sub BleedingTimer(Index As Integer)
'> Bleeding
    If Char(Index).Bleeding > 0 Then
        If Char(Index).HP - Char(Index).Bleeding > 0 Then
            Char(Index).HP = Char(Index).HP - Char(Index).Bleeding
            Send Index, "You " & bRED & "bleed" & WHITE & "!" & vbCrLf & "HP: " & Char(Index).HP
            TransmitLocal Index, Char(Index).Name & bRED & " bleeds" & WHITE & "!"
        Else
            Send Index, "You lost in the battle and would most likely die if it wasn't ALPHA test of BeMUD."
            TransmitLocal Index, Char(Index).Name & " dies from " & HeShe(Char(Index).Gender, "HisHer") & " wounds"
        End If
    End If
'< Bleeding
End Sub
Sub AddDelay(Index As Integer, Seconds As Integer, CommandName As String, _
            Optional Mnum = 0, Optional NameIndex As Integer = 0)
With Char(Index)
    
    .Delay.Busy = True
    .Delay.Command = CommandName
    .Delay.PCTarget = NameIndex
    .Delay.MobItemVnum = Mnum
    
    frmMain.tmrDelayedComs(Index).Enabled = True
    frmMain.tmrDelayedComs(Index).Interval = Seconds * IntervalFormat

End With
End Sub
Sub RemoveDelay(Index As Integer)
With Char(Index)
    
    frmMain.tmrDelayedComs(Index).Enabled = False
    .Delay.PCTarget = 0
    .Delay.MobItemVnum = Empty
    .Delay.Busy = False

End With
End Sub
Function TargetChanged(Index%, NameIndex%, Mnum%) As Boolean
Dim TargetNotHere As Boolean

TargetChanged = False

'\B/------------------------Checking the target is still here------------------------
        If CheckMPI(Mnum) Then If InStr(QteMe(Mobs(Index)), Mnum) = 0 Then TargetNotHere = True
        If CheckMPI(NameIndex) And InStr(QteMe(PCs(Index)), QteMe(NameIndex)) = 0 Then TargetNotHere = True
'/E\------------------------Checking the target is still here------------------------
            
'\B/---------------------------------Target is gone---------------------------------
        If TargetNotHere Then
            Send Index, "He fled!"
            RemoveDelay Index
            TargetChanged = True
        End If
'/E\---------------------------------Target is gone---------------------------------
End Function

