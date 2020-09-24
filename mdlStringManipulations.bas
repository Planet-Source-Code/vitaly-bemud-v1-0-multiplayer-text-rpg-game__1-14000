Attribute VB_Name = "mdlStringManipulations"
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
Public Function Proper(ByVal Str As String) As String
    '[Description]
    'Rudimentary routine to convert a mixed
    '     case string to Sentence case
    '[Declarations]
    Dim flgNextUpper As Boolean 'Is next alpha character the start of a new sentence
    Dim intIndex As Integer 'Current character being tested
    '[Code]
    Str = LCase(Str)
    flgNextUpper = True
    For intIndex = 1 To Len(Str)
        If Mid(Str, intIndex, 1) >= "a" _
          And Mid(Str, intIndex, 1) <= "z" _
          And flgNextUpper Then
        'Convert the current character
            Mid(Str, intIndex, 1) = UCase(Mid(Str, intIndex, 1))
        flgNextUpper = False
        End If
    If InStr(".!?:", Mid(Str, intIndex, 1)) Then
        'End of sentence reached
        flgNextUpper = True
    End If
Next 'character In String
Proper = Str
End Function
Function GetWordByNum(lWordNum As Integer, ByVal sSentence As String, Optional Devider As String = ",") As String
'The function gets a sentence and returns the word number lWordNum
Dim Temp() As String
    Temp = Split(sSentence, Devider)
    If UBound(Temp) >= lWordNum - 1 And LBound(Temp) <= lWordNum - 1 Then GetWordByNum = Temp(lWordNum - 1)
'\B/--------------------For Visual Basic 5------------------------------
'Dim I%, LastSpace%
'    LastSpace = 1
'    Word = " " & Word & " "
'    For I = 1 To Number
'        GetWord = Mid(Word, InStr(LastSpace, Word, " ") + 1, InStr(LastSpace + 1, Word, " ") - InStr(LastSpace, Word, " ") - 1)
'        LastSpace = InStr(LastSpace + 1, Word, " ")
'    Next I
'/E\--------------------For Visual Basic 5------------------------------
End Function
Function GetNumByWord(ByVal sWord As String, ByVal sSentence As String, Optional Devider As String = ",") As Integer
'The function gets a sentence and returns the number of the word sWord, returns -1 if not found.
    Dim I%, Temp$()
    GetNumByWord = -1
    Temp = Split(sSentence, Devider)
    For I = 0 To UBound(Temp)
        If Temp(I) = sWord Then GetNumByWord = I + 1: Exit For
    Next I
'\B/--------------------For Visual Basic 5------------------------------
'    Dim WordLocation%, CommaLocation%
'    sSentence = Devider & sSentence & Devider
'    sWord = Devider & sWord & Devider
'    WordLocation = InStr(sSentence, sWord)
'    Do
'        I = I + 1
'        CommaLocation = InStr(CommaLocation + 1, sSentence, ",")
'        If CommaLocation = WordLocation Then GetNumByWord = I: Exit Function
'    Loop Until InStr(sSentence, ",") = 0
'/E\--------------------For Visual Basic 5------------------------------
End Function
'Function GetWordsByNums(StartNum As Integer, EndNum As Integer, ByVal sSentence As String, Optional Devider As String = ",")
''This function gets a few words from a sentance by having starting and ending nums
'Dim Temp$
'    Temp = Split(sSentence, Devider)
'    For I = StartNum To EndNum
'End Function
Function StringToArray(ByVal sString As String, Optional Devider As String = ",") As Variant
    StringToArray = Split(sString, Devider)
'\B/--------------------For Visual Basic 5------------------------------
'Dim I%, Arr()
'    Do Until InStr(sString, Devider) = 0
'        I = I + 1
'        ReDim Arr(I)
'        sString = Mid(sString, InStr(sString, Devider) + 1)
'        Arr(I) = Left(sString, IIf(InStr(sString, Devider), InStr(sString, Devider) - 1, Len(sString)))
'    Loop
'    If I > 0 Then StringToArray = Arr Else ReDim Arr(0)
'/E\--------------------For Visual Basic 5------------------------------
End Function
Function CountWords(sString As String, Optional Devider As String = " ") As Integer
    'Counts the number of the words in the string (Sorry Vb5, use above functions to recreate this)
    CountWords = UBound(Split(sString, Devider)) + 1
End Function
Function CountChars(sString As String, Char As String) As Integer
Dim I%, Start%
    Start = 1
    Do
        If InStr(Start, sString, Char) > 0 Then I = I + 1: Start = 1 + InStr(Start, sString, Char)
    Loop Until InStr(Start, sString, Char) = 0
    CountChars = I
End Function
Function LettersOnly(ByVal sString As String) As Boolean
    LettersOnly = True
    Dim I As Integer
    sString = LCase(sString)
    For I = 1 To Len(sString)
        If (Asc(Mid(sString, I, 1)) < 97 Or Asc(Mid(sString, I, 1)) > 122) And Asc(Mid(sString, I, 1)) <> 32 Then LettersOnly = False: Exit Function
    Next I
End Function
Function QteMe(ByVal Str As Variant, Optional Devider$ = ",") As String
'\B/---------------------------------"Me" to ",Me,"---------------------------------
QteMe = Devider$ & Str & Devider$
'/E\---------------------------------"Me" to ",Me,"---------------------------------
End Function
Function AddToString(ByVal sString$, ByVal WordToAdd$, Optional Devider$ = ",")
    If Len(sString) > 0 And Len(WordToAdd) > 0 Then
        AddToString = WordToAdd$ & Devider$ & sString
        ElseIf Len(WordToAdd) > 0 Then AddToString = WordToAdd$
        Else: AddToString = sString
    End If
End Function
Function RemoveFromString(ByVal sString$, ByVal WordToRemove$, Optional Devider$ = ",", Optional HowMany% = -1)
    If sString = WordToRemove$ Then RemoveFromString = "": Exit Function
    sString = Replace(Devider$ & sString & Devider$, Devider & WordToRemove & Devider, Devider, , HowMany)
    RemoveFromString = Mid(sString, 2, Len(sString) - 2) 'Removing the quotes
End Function
Function LongInStr(Sentence$, SearchWord$, WhichOne%) As Integer
Dim Start%, I%
    For I = 1 To WhichOne
        Start = InStr(Start + 1, Sentence, SearchWord)
    Next I
        LongInStr = Start
End Function
