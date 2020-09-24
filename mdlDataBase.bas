Attribute VB_Name = "mdlDataBase"
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
Public DB As Database 'The var that will hold the whole BeMud.mdb file
Private Type RecordType
    SnapShot As Recordset
    Dynaset As Recordset
    Table As String
End Type
Public CharRecord As RecordType
Public Function dbIsExist(RS As Recordset, ByVal Table As String, Index As Integer, FindWhere As String, FindWhat As String) As Boolean
'Function checks if in the Field there is a SearchString using a SQL query.
'Returns True/False.
    dbIsExist = False
    Set RS = DB.OpenRecordset("Select * from " & Table & " where " & FindWhere & "=" & QteMe(FindWhat, QT))
    If RS.RecordCount > 0 Then dbIsExist = True
End Function
Public Function DbFind(RS As Recordset, ByVal Table As String, Index As Integer, FindWhere As String, FindWhat As String, ReturnWhat As String) As String
    Set RS = DB.OpenRecordset("Select * from " & Table & " where " & FindWhere & "=" & QteMe(FindWhat, QT))
    If Not RS.EOF And Not RS.BOF Then _
     DbFind = IIf(IsNull(RS(ReturnWhat)), "", RS.Fields(ReturnWhat))
End Function
Public Sub DbEdit(ByVal RS As Recordset, ByVal Table As String, Index As Integer, FindWhere As String, FindWhat As String, Field As String, NewData As String)
'This sub edits ONE field
    Set RS = DB.OpenRecordset("Select * from " & Table & " where " & FindWhere & "=" & QteMe(FindWhat, QT))
    'RS.FindFirst FindWhere & "=" & Chr(34) & FindWhat & Chr(34)
    If Not RS.EOF And Not RS.BOF Then
       RS.Edit
           RS.Fields(Field) = NewData
       RS.Update
    End If
End Sub


Sub DBHelp()
    Dim DB As Database
    'This is the object that will hold the connection
    'to our database
    Dim Record As Recordset
    'This is the object that will hold a set of
    'records coming back from the database
    Set DB = OpenDatabase("D:\Desktop\My documents\1 Internet\2 Received Files - ICQ\Norhat\bemud.mdb")
    Set Record = DB.OpenRecordset("Select DBRef, Name from Characters where [Eq head]=0", dbOpenDynaset)
    Record.MoveFirst
    Record.Seek
    Record.FindFirst "Name=" & QT & "NewGuy" & QT
    Record.Edit
    Record.Fields(1) = "Vitaly"
    Record.Update
    Record.MoveLast
    Record.Delete
End Sub
