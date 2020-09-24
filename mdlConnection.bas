Attribute VB_Name = "mdlConnection"
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
Sub DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim PartData As String
    frmMain.wskAccept(Index).GetData PartData
    'If recieved the backpsace char from Telnet
    If Asc(PartData) = 8 And Len(Char(Index).Data) > 0 Then
        Char(Index).Data = Left(Char(Index).Data, Len(Char(Index).Data) - 1)
        Exit Sub
    End If
    Char(Index).Data = Char(Index).Data & PartData
    If InStr(Char(Index).Data, vbCrLf) Or InStr(Char(Index).Data, vbLf) Then
        Dim UserInput As Variant, I As Integer
        Char(Index).Data = Replace(Char(Index).Data, vbCrLf, vbCr)
        Char(Index).Data = Replace(Char(Index).Data, vbLf, vbCr)
        Char(Index).Data = Left(Char(Index).Data, Len(Char(Index).Data) - 1)
'\B/------------------Splits the input and make it ready to process------------------
        UserInput = Split(IIf(Char(Index).Data = "", vbLf, Char(Index).Data), vbCr) '
      For I = LBound(UserInput) To UBound(UserInput)
'\B/-----------------------------------Spy command-----------------------------------
        If Len(PlayerList.Spy) > 0 Then
            Dim Arr, AdminUser%
            Arr = StringToArray(PlayerList.Spy)
            For AdminUser = LBound(Arr) To UBound(Arr)
                If Index <> Arr(AdminUser) Then _
                Send Arr(AdminUser), "[SPY] " & Char(Index).Name & ": " & Char(Index).Data
            Next AdminUser
        End If
'/E\-----------------------------------Spy command-----------------------------------
'\B/------------------Checking to what Stage the command was sent-----------------------------------------------------------------------------
'Stages: 3 - Game mode, 1 - Creation mode
        If UserInput(I) = vbLf Then UserInput(I) = ""
        Select Case Char(Index).GameState
        Case "Game" 'In game
            If UserInput(I) <> "" Then DoCommands Index, CStr(UserInput(I))
        Case "Gender", "Name", "NameConfirm", "PasswordCheck", "PasswordChoosing", "PasswordConfirm", "Race"
            If UserInput(I) <> "" Then DoCreation Index, CStr(UserInput(I))
        Case "Qued"
            If UserInput(I) <> "q" Then Send Index, Char(Index).QdText, ToWrap:=False, CheckForSplit:=True _
              Else: Char(Index).GameState = "Game"
        End Select
        'Debug.Print "Data sent: " & Timer
'\B/-------------------------------------Prompt-------------------------------------
        If UBound(Char) >= Index Then _
          If Char(Index).GameState = "Game" Then Send Index, "> ", RET:=""
'/E\-------------------------------------Prompt-------------------------------------
'/E\------------------Checking to what Stage the command was sent-----------------------------------------------------------------------------
      Next I
'/E\------------------Splits the input and make it ready to process------------------
    Char(Index).Data = ""
    End If
End Sub
Sub ConnectionRequest(ByVal requestID As Long)
Dim LastConnection As Integer
    If Len(PlayerList.FreeIndex) > 0 Then
        LastConnection = GetFreeVNum(PlayerList.FreeIndex)
    Else
        LastConnection = frmMain.wskAccept.UBound + 1
        Load frmMain.wskAccept(LastConnection)
        Load frmMain.tmrBleeding(LastConnection)
        Load frmMain.tmrDelayedComs(LastConnection)
    End If
    If LastConnection > UBound(Char) Then ReDim Preserve Char(1 To LastConnection)
    frmMain.wskAccept(LastConnection).Close
    frmMain.Log "Accepting access on " & LastConnection & " Winsock.", Now & " Server"
    frmMain.wskAccept(LastConnection).Accept requestID
    
    Char(LastConnection).GameState = "Name"
    Char(LastConnection).Name = "Unknown"
    Send LastConnection, CLRSCR & RET & bRED & GetIni("Logo", "Draw", "Graphics") & WHITE & RET & RET & "Welcome to BeMUD, please enter your name: ", RET:="", ToWrap:=False
End Sub
    'Closing connection and resetting the personal Vars
Sub CloseConnection(Index As Integer, Optional QuitMsg$ = "yawns and goes to sleep.")
 Dim I As Integer
    Call RemoveAll(Index) 'Remvoe everything weared
'\B/-----------------------------Drop all the equipment-----------------------------
    If Char(Index).Items <> "" Then
        Area(Char(Index).Area).Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ).Items = _
          AddToString(Area(Char(Index).Area).Room(Char(Index).locX, Char(Index).locY, Char(Index).locZ).Items, Char(Index).Items)
        Char(Index).Items = RemoveFromString(Char(Index).Items, Char(Index).Items)
    End If
'/E\-----------------------------Drop all the equipment-----------------------------
    
    If Char(Index).GameState = "Game" Then
        TransmitLocal Index, Char(Index).Name & " " & QuitMsg
        RemovePC Index
    End If
    
    Select Case Char(Index).Status
    Case Admin
        PlayerList.Admins = RemoveFromString(PlayerList.Admins, Trim(Index))
    Case Immortal
        PlayerList.Immortals = RemoveFromString(PlayerList.Immortals, Trim(Index))
    Case Mortal
        PlayerList.Mortals = RemoveFromString(PlayerList.Mortals, Trim(Index))
    End Select
        
    frmMain.Log "Closing access on " & Index & " Winsock.", Now & " Server"
    
    If Index = frmMain.wskAccept.UBound And Index > 1 Then
        Unload frmMain.wskAccept(Index): Unload frmMain.tmrBleeding(Index)
        Unload frmMain.tmrDelayedComs(Index)
    Else
        frmMain.wskAccept(Index).Close
        frmMain.tmrBleeding(Index).Enabled = False
        frmMain.tmrDelayedComs(Index).Enabled = False
    End If
    
    If frmMain.wskAccept.UBound < UBound(Char) Then
        ReDim Preserve Char(frmMain.wskAccept.UBound)
    Else
        Char(Index).Data = ""
        Char(Index).GameState = ""
        Char(Index).locX = 0: Char(Index).locY = 0: Char(Index).locZ = 666 '666, so he won't be thought as someone in 1,1,1 room
        Char(Index).Name = ""
        Char(Index).Spy = False
        Char(Index).Bleeding = 0
        PlayerList.FreeIndex = AddToString(PlayerList.FreeIndex, Index) 'Adds the free vnum to the list
    End If
    
End Sub
Sub WinsockError(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print "Error " & Number & " - " & Description
    If Number = 10054 Then
    Call CloseConnection(Index)
    End If
End Sub

Function Wrapped(ByVal Data As String, Optional CurrentColor As String = "", _
         Optional RETRN As String = vbCrLf, Optional Indent As String = "") As String

Dim Wrap%
    'If Color <> 7 Then CurrentColor = Ansi(Color) Else CurrentColor = Chr(27) & "[0m" & Chr(27) & "[37m"
    On Error Resume Next
    Wrap = 60
    If Len(Data) < Wrap Then
        Wrapped = CurrentColor & Data & RETRN
    'Wrapping the text if it is too long.
    Else
        Dim Enters%
        Do Until InStr(Wrap, Data, " ") = 0
            'Checks for RETRN in the string and wraps it by RETRNs and Spaces.
            If InStr(Data, vbCrLf) And InStr(Data, vbCrLf) < InStr(Wrap, Data, " ") Then
                Wrapped = Wrapped & CurrentColor & Left(Data, InStr(Data, vbCrLf))
                Data = Mid(Data, InStr(Data, vbCrLf) + 1)
            Else
                Wrapped = Wrapped & CurrentColor & Left(Data, InStr(Wrap, Data, " ") - 1) & vbCrLf
                Data = Indent & Mid(Data, InStr(Wrap, Data, " ") + 1)
            End If
        Loop
        'debug.Print CountChars(Wrapped, vbCrLf)
        Wrapped = Wrapped & CurrentColor & Data & RET
    End If
End Function
'Sending the data to the Index
Sub Send(ByVal Index As Integer, ByVal Data As String, Optional Color As String = "", _
    Optional ToWrap As Boolean = True, Optional RET As String = vbCrLf, _
    Optional CheckForSplit As Boolean, Optional Indent As String = "")
    
    On Error Resume Next
    If ToWrap Then Data = Wrapped(Data, Color, RET, Indent)
'\B/-----------------------------Press enter to continue-----------------------------
        If CheckForSplit Then
            If CountChars(Data, vbCrLf) > 20 Then
                Dim Interrupt%
                Interrupt = LongInStr(Data, vbCrLf, 20)
                Char(Index).QdText = Mid(Data, Interrupt + 2)
                Data = Mid(Data, 1, Interrupt)
                Data = Data & RET & "Press anything to continue, 'q' to stop." & RET
                Char(Index).GameState = "Qued"
            Else
                Char(Index).GameState = "Game"
            End If
        End If
'/E\-----------------------------Press enter to continue-----------------------------
        frmMain.wskAccept(Index).SendData Data
End Sub
'Transmits the message to people in the same room with the the transmitor
Sub TransmitLocal(ByVal Index As Integer, ByVal Transmit As String, Optional NoSend As Integer = -1, _
    Optional Color As String = "")
    
    Dim I%, WrappedText$
    Dim Arr As Variant
    Arr = StringToArray(PCs(Index))
    If UBound(Arr) > 0 Then
        WrappedText = Wrapped(Transmit, Color)
        For I = LBound(Arr) To UBound(Arr)
            If Arr(I) <> Index And Arr(I) <> NoSend Then Send Val(Arr(I)), WrappedText, , ToWrap:=False
        Next I
    End If
    End Sub
'Transmits the message to a list of people
Sub TransmitList(ByVal List As String, ByVal Transmit As String, Optional NoSend As Integer = -1, _
    Optional Color As String = "")
    
    Dim I%, WrappedText$
    Dim Arr As Variant
    Arr = StringToArray(List)
    If UBound(Arr) >= 0 Then
        WrappedText = Wrapped(Transmit, Color)
        For I = LBound(Arr) To UBound(Arr)
            If Arr(I) <> NoSend Then Send Val(Arr(I)), WrappedText, , ToWrap:=False
        Next I
    End If
End Sub
