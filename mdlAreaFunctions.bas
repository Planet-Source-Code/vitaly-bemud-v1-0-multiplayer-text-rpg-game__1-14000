Attribute VB_Name = "mdlArea"
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
'> Area vars (duh)
    Public Type RoomVars
        Description As String 'That's each char unique input string
        PCs As String
        Exits As String
        Items As String
        Mobs As String
    End Type
    Public Type AreaVars
        Room() As RoomVars
        Name As String 'Area name
    End Type
    Public Area() As AreaVars
'< Area vars

'Gets the description of the current room
Function CurrentDesc(Index As Integer) As String
    CurrentDesc = Area(Char(Index).Area) _
      .Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ).Description
End Function
'Gets the aviable exits of the current room
Function CurrentExits(Index As Integer) As String
    CurrentExits = LCase(Area(Char(Index).Area) _
      .Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ).Exits)
End Function
'Adds the PC to the room records
Sub AddPC(ByVal Index)
    'Sets the value of the PCs to a temp variable for a more readable code
    With Area(Char(Index).Area).Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ)
    
    'and comma if it is not, so the result should be like 6,3,7,4
    If Len(.PCs) > 0 Then _
      .PCs = Index & "," & .PCs _
      Else: .PCs = Index 'For VB5 "," & Index ONLY (No if)
    
    End With
End Sub
'Removes the PC to the room records
Sub RemovePC(ByVal Index)
    
    With Area(Char(Index).Area).Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ)

    If Trim(Index) = .PCs Then .PCs = ""
    .PCs = Replace(QteMe(.PCs), QteMe(Index), ",")
    .PCs = Mid(.PCs, 2, Len(.PCs) - 2)

    End With
End Sub
'Returns the Pcs in the room (Not parsed)
Function PCs(Optional Index% = 0, Optional Mnum% = 0) As Variant

If CheckMPI(Index) Then PCs = Area(Char(Index).Area) _
      .Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ).PCs

If CheckMPI(Mnum) Then PCs = Area(Char(Mnum).Area) _
      .Room(Char(Mnum).locX, Char(Mnum).locY, Char(Mnum).locZ).PCs

End Function
Function Items(Optional Index% = 0, Optional Mnum% = 0) As Variant

If CheckMPI(Index) Then Items = Area(Char(Index).Area) _
      .Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ).Items

If CheckMPI(Mnum) Then Items = Area(Char(Mnum).Area) _
      .Room(Char(Mnum).locX, Char(Mnum).locY, Char(Mnum).locZ).Items

End Function
Function Mobs(Optional Index% = 0, Optional Mnum% = 0) As String

If CheckMPI(Index) Then Mobs = Area(Char(Index).Area) _
      .Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ).Mobs

If CheckMPI(Mnum) Then Mobs = Area(Char(Mnum).Area) _
      .Room(Char(Mnum).locX, Char(Mnum).locY, Char(Mnum).locZ).Mobs

End Function
'Adds the Item to the room records
Sub AddItem(Index%, Inum%)
    
    With Area(Char(Index).Area).Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ)
    
    'For result like 3 5,4 4,7 3
    If Len(.Items) > 0 Then .Items = Inum & "," & .Items _
      Else: .Items = Inum
    End With
End Sub
Sub RemoveItem(Index%, Inum%)
    
    With Area(Char(Index).Area).Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ)

    If Trim(Inum) = .Items Then .Items = ""
    .Items = Replace(QteMe(.Items), QteMe(Inum), ",")
    .Items = Mid(.Items, 2, Len(.Items) - 2)

    End With
End Sub
Sub AddMob(Index%, Mnum%)
    With Area(Char(Index).Area).Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ)
    
    'Little messy :( Adds Char index only if it is empty and Char index
    'and comma if it is not, so the result should be like 6,3,7,4
    If Len(.Mobs) > 0 Then _
      .Mobs = Mnum & "," & .Mobs _
      Else: .Mobs = Mnum
      
    End With
End Sub
Sub RemoveMob(Mnum%)
    
    With Area(Mob(Mnum%).Area).Room(Mob(Mnum%).locX, Mob(Mnum%).locY, Mob(Mnum%).locZ)

    If Trim(Mnum%) = .Mobs Then .Mobs = "": Exit Sub
    .Mobs = Replace(QteMe(.Mobs), QteMe(Mnum), ",")
    .Mobs = Mid(.Mobs, 2, Len(.Mobs) - 2)
    
    End With
End Sub
Function SearchAreaName(ArrayName() As AreaVars, LookFor As String) As Integer
    Dim I%
    SearchAreaName = 0
    For I = LBound(ArrayName) To UBound(ArrayName)
        Rem VNEW: ArrayName(I).ID ---> ArrayName(I).Name
        If LCase(ArrayName(I).Name) = LCase(LookFor) Then SearchAreaName = I: Exit Function
    Next I
End Function
