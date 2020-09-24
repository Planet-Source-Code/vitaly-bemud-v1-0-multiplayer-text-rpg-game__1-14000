Attribute VB_Name = "mdlGlobalProcedures"
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
Option Base 1
    Public GDic As New Dictionary
    Public GRecord As Recordset 'System RecordSet
    Public ForbiddenNames As String 'Character names that can't be used
'> Emote vars
    Public Type EmoteVars
        ID As String
        Self As String
        Others As String
        SelfTarget As String
        OthersTarget As String
        Target As String
    End Type
    Public Emotes() As EmoteVars
'< Emote vars
'\B/-----------------------------------Time check-----------------------------------
Public Declare Function GetTickCount Lib "kernel32" () As Long
'/E\-----------------------------------Time check-----------------------------------

'\B/----------------------------------Always on top----------------------------------
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    Public Const SWP_NOMOVE = 2
    Public Const SWP_NOSIZE = 1
    Public Const WndFlags = SWP_NOMOVE Or SWP_NOSIZE
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2
'/E\----------------------------------Always on top----------------------------------

'\B/------------------------------------Constants------------------------------------
Public Const RET As String = vbCrLf 'Enter
Public Const QT As String = """" 'Const for double quote
Public Const IntervalFormat As Integer = 1000 'Const that counts the interval on AddDelay subs
'/E\------------------------------------Constants------------------------------------

'\B/------------------------------------Ini file------------------------------------
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'/E\------------------------------------Ini file------------------------------------

'\B/-----------------------------------Ansi codes-----------------------------------
Public Const CLRSCR As String = "[2J" 'Clears the screen

Public Const BLACK As String = "[0m[30m"
Public Const RED As String = "[0m[31m"
Public Const GREEN As String = "[0m[32m"
Public Const YELLOW As String = "[0m[33m"
Public Const BLUE As String = "[0m[34m"
Public Const MAGNETA As String = "[0m[35m"
Public Const LIGHTBLUE As String = "[0m[36m"
Public Const WHITE As String = "[0m[37m"

Public Const bBLACK As String = "[1m[30m"
Public Const bRED As String = "[1m[31m"
Public Const bGREEN As String = "[1m[32m"
Public Const bYELLOW As String = "[1m[33m"
Public Const bBLUE As String = "[1m[34m"
Public Const bMAGNETA As String = "[1m[35m"
Public Const bLIGHTBLUE As String = "[1m[36m"
Public Const bWHITE As String = "[1m[37m"
'/E\-----------------------------------Ansi codes-----------------------------------

Function CheckMPI(MPI As Variant) As Boolean
'The function checks if the variable is an array (Mob/Item is here) or above 0 (Pc is here).
    If MPI > 0 Then CheckMPI = True: Exit Function
End Function
'Function that checks if what index of array is equal to LookFor
Function SearchEmoteID(ArrayName() As EmoteVars, LookFor As String) As Integer
    Dim I%
    SearchEmoteID = 0
    For I = LBound(ArrayName) To UBound(ArrayName)
        If LCase(ArrayName(I).ID) = LookFor Then SearchEmoteID = I: Exit Function
    Next I
End Function
Function GetCondition(iCondition%, MaxHP%) As String
    Select Case iCondition
        Case MaxHP%
            GetCondition = "healthy"
        Case Is > MaxHP% / 100 * 90
            GetCondition = "slightly scratched"
        Case Is > MaxHP% / 100 * 80
            GetCondition = "scratched"
        Case Is > MaxHP% / 100 * 70
            GetCondition = "bruised"
        Case Is > MaxHP% / 100 * 60
            GetCondition = "slightly wounded"
        Case Is > MaxHP% / 100 * 50
            GetCondition = "badly wounded"
        Case Is > MaxHP% / 100 * 40
            GetCondition = "bleeding"
        Case Is > MaxHP% / 100 * 30
            GetCondition = "bleeding freely"
        Case Is > MaxHP% / 100 * 20
            GetCondition = "gushing blood"
        Case Is > MaxHP% / 100 * 10
            GetCondition = "torn to flesh"
        Case Else
            GetCondition = "too wounded to function"
    End Select
End Function
Function GetFreeVNum(VNums As String) As Integer
Dim Arr
    Arr = Split(VNums, ",")
    GetFreeVNum = Arr(UBound(Arr))
    VNums = RemoveFromString(VNums, Arr(UBound(Arr)))
End Function
Function GetIni(Class As String, VarToGet As String, IniFileName As String) As String
Dim RET As Long
Dim Temp As String * 2050
Dim lpAppName As String, lpKeyName As String, lpDefault As String, lpFileName As String

IniFileName = App.Path & "\" & IniFileName & ".ini"
lpAppName = Class 'Class name ( [User] )
lpDefault = IniFileName 'Ini location and file name
lpFileName = IniFileName 'Ini location and file name
RET = GetPrivateProfileString(lpAppName, VarToGet, IniFileName, Temp, Len(Temp), IniFileName)
GetIni = Mid(Temp, 1, RET)
If GetIni = IniFileName Then GetIni = "Error" Else _
  GetIni = Replace(GetIni, "<r>", vbCrLf)
End Function
Sub PutIni(Class As String, VarToPut As String, Value As String, IniFileName As String)
    Dim lpAppName As String, lpFileName As String, lpKeyName As String, lpString As String
    Dim RET As Long
    lpAppName = Class 'Class name ( [User] )
    lpKeyName = VarToPut 'Variable
    lpString = Value 'Value
    lpFileName = App.Path & "\" & IniFileName & ".ini" 'Ini location and file name
    RET = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
    If RET = 0 Then
      Beep
    End If
End Sub
'Function that checks if something is a part of an array
Function IsMember(ArrayName() As String, LookFor As String) As Boolean
'Entrance: ArrayName to look in and String to look for
    Dim Arr$(), I%
    Arr = Filter(ArrayName, LookFor)
    If UBound(Arr) >= 0 And Len(Arr(0)) > 0 Then IsMember = True Else IsMember = False
'Exit: True if found, false if not.
End Function
'Function that checks if what index of array is equal to LookFor
Function WhatMember(ArrayName() As String, LookFor As String) As Integer
    Dim Arr$(), I%
    WhatMember = 0
    For I = LBound(ArrayName) To UBound(ArrayName)
        If LCase(ArrayName(I)) = LookFor Then WhatMember = I: Exit Function
    Next I
End Function
'Increase sub
Sub Inc(Num As Integer)
    Num = Num + 1
End Sub
'Decrease sub
Sub Dec(Num As Integer)
    Num = Num - 1
End Sub

Function Ansi(Color As Integer) As String
    Select Case Color
    Case Is <= 7
        Ansi = Chr(27) & "[0m" & Chr(27) & "[3" & Color & "m"
    Case Is >= 8
        Ansi = Chr(27) & "[1m" & Chr(27) & "[3" & Color - 8 & "m"
    End Select
'ANSI Name  Code Sequence   Action
'
'*******************************************************************************
'
'     ED    ESC[ Pn J       Erases all or part of a display.
'                           Pn=0: erases from active position to end of display.
'                           Pn=1: erases from the beginning of display to
'                                 active position.
'                           Pn=2: erases entire display.
'
'     EL    ESC[ Pn K       Erases all or part of a line.
'                           Pn=0: erases from active position to end of line.
'                           Pn=1: erases from beginning of line to active
'                                 position.
'                           Pn=2: erases entire line.
'
'     ECH   ESC[ Pn X       Erases Pn characters.
'
'     CBT   ESC[ Pn Z       Moves active position back Pn tab stops.
'
'     SU    ESC[ Pn S       Scroll screen up Pn lines, introducing new blank
'                           lines at bottom.
'
'     SD    ESC[ Pn T       Scroll screen down Pn lines, introducing new blank
'                           lines at top.
'
'     CUP   ESC[ P1;P2 H    Moves cursor to location P1 (vertical)
'                           and P2 (horizontal).
'
'     HVP   ESC[ P1;P2 f    Moves cursor to location P1 (vertical)
'                           and P2 (horizontal).
'
'     CUU   ESC[ Pn A       Moves cursor up Pn number of lines.
'
'     CUD   ESC[ Pn B       Moves cursor down Pn number of lines.
'
'     CUF   ESC[ Pn C       Moves cursor Pn spaces to the right.
'
'     CUB   ESC[ Pn D       Moves cursor Pn spaces backward.
'
'     HPA   ESC[ Pn '       Moves cursor to column given by Pn.
'
'     HPR   ESC[ Pn a       Moves cursor Pn characters to the right.
'
'     VPA   ESC[ Pn d       Moves cursor to line given by Pn.
'
'     VPR   ESC[ Pn e       Moves cursor down Pn number of lines.
'
'     IL    ESC[ Pn L       Inserts Pn new, blank lines.
'
'     ICH   ESC[ Pn @       Inserts Pn blank places for Pn characters.
'
'     DL    ESC[ Pn M       Deletes Pn lines.
'
'     DCH   ESC[ Pn P       Deletes Pn number of characters.
'
'     CPL   ESC[ Pn F       Moves cursor to beginning of line, Pn lines up.
'
'     CNL   ESC[ Pn E       Moves cursor to beginning of line, Pn lines down.
'
'     SGR   ESC[ Pn m       Changes display mode.
'                           Pn=0: Resets bold, blink, blank, underscore, and
'                           reverse.
'                           Pn=1: Sets bold (light_color).
'                           Pn=4: Sets underscore.
'                           Pn=5: Sets blink.
'                           Pn=7: Sets reverse video.
'                           Pn=8: Sets blank (no display).
'                           Pn=10: Select primary font.
'                           Pn=11: Select first alternate font.
'                           Pn=12: Select second alternate font.
'
'           ESC[ 2h         Lock keyboard. Ignores keyboard input until
'                           unlocked.
'
'           ESC[ 2i         Send screen to host.
'
'           ESC[ 2l         Unlock keyboard.
'
'           ESC[ 3 C m      Selects foreground colour C.
'
'           ESC[ 4 C m      Selects background colour C.
'
'                           C=0  Black
'                           C=1  Red
'                           C=2  Green
'                           C=3  Yellow
'                           C=4  Dark Blue
'                           C=5  Magenta
'                           C=6  Light Blue
'                           C=7  White
'
'****************************************************************************
'
'Here are the IBM / Ansi Characters
End Function
Function GlobalOpenTags(ByVal MainStr$, Optional ByVal Index%, Optional ByVal Mnum%, Optional BodyPartName As String) As String
Dim Str As String
    Str = MainStr
    
'\B/-----------------------------Opening character tags-----------------------------
    If CheckMPI(Index) Then
    
        Str = Replace(Str, "<targetname>", Char(Index).Name)
        
    End If
'/E\-----------------------------Opening character tags-----------------------------
'\B/--------------------------------Opening mob tags--------------------------------
    If CheckMPI(Mnum) Then
    
        Str = Replace(Str, "<targetname>", Mob(Mnum).Name)
    
    End If
'/E\--------------------------------Opening mob tags--------------------------------
        Str = Replace(Str, "<bodypartname>", BodyPartName)
    
    GlobalOpenTags = Str
End Function
