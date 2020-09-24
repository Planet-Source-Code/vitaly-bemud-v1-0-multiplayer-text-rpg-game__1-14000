Attribute VB_Name = "mdlItem"
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
'> Item vars
    Public Type ItemVars
        ID As Integer 'Unique ID of item type
        Name As String
        Aliases As String
        Description As String
        Type As String
        Subtype As String
        Wear As String
        AC As Integer
        Damage As Integer
    End Type
    Public PrototypeItem() As ItemVars 'Keeps and stores the prototype of each mob type (ID)
    Public Item() As ItemVars
    Public IFreeVNums As String 'Has the VNums of items that got free (e.g If item vanished)
'< Item vars
'This function is an annoying one. Basically it sorts the weared items by importance,
'It is used in LOOK and EQUIPMENT commands
Function SortItemsList(ItemsList$) As Variant
Dim WearsItemsArr As Variant, ItemsValue$(1 To 5)
Dim Vnum%, ValueLevel%, I%
    WearsItemsArr = StringToArray(ItemsList)
    'Values:
    'Value 1: Value for items holded in phand
    'Value 2: Value for items worn on body
    For I = LBound(WearsItemsArr) To UBound(WearsItemsArr)
        Vnum = GetWordByNum(2, WearsItemsArr(I), " ")
        Select Case Item(Vnum).Wear
        Case "phand"
            ValueLevel = 1
            ItemsValue(ValueLevel) = Vnum
        Case "torso"
            ValueLevel = 2
            ItemsValue(ValueLevel) = Vnum
        End Select
    Next I
    SortItemsList = ItemsValue
End Function
Function WearsItems(ItemsValue As Variant, PHandName$, Gender$)
    '\B/----------------------------------------See wearing equipment----------------------------------------
    Dim Noun$, Verb$, Proposition$ 'These are used to construct the final look
    Dim I%
        For I = 1 To UBound(ItemsValue)
            If ItemsValue(I) <> "" Then
                Select Case Item(ItemsValue(I)).Type
                Case "armor"
                    Verb = " wears "
                    Proposition = " on "
                Case "weapon"
                    Verb = " holds "
                    Proposition = " in "
                End Select
                Select Case Item(ItemsValue(I)).Wear
                Case "phand"
                    Noun = " " & PHandName & ". "
                Case "torso"
                    Noun = " torso. "
                End Select
                WearsItems = WearsItems & _
                  StrConv(HeShe(LCase(Gender), "HeShe"), vbProperCase) & Verb & _
                  Item(ItemsValue(I)).Name & Proposition & HeShe(LCase(Gender), "HisHer") & Noun
            End If
        Next I
    '/E\----------------------------------------See wearing equipment----------------------------------------
End Function
Function ActivateItem(ID As Integer) As Integer
'This function creates the item with the item in the Item() array and returns its vnum
    Dim FreeVnum As Integer
    If Len(IFreeVNums) > 0 Then
        FreeVnum = GetFreeVNum(IFreeVNums)
    ElseIf Item(UBound(Item)).ID > 0 Then
        FreeVnum = UBound(Item) + 1
        ReDim Preserve Item(1 To FreeVnum)
    Else
        FreeVnum = 1
    End If
    
    Item(FreeVnum) = PrototypeItem(ID)
    ActivateItem = FreeVnum
End Function
