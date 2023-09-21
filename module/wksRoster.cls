VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wksRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim hasChanged As Boolean

Private Sub Worksheet_Activate()

    Call CreateOrDisplayNextMTG(wksHome.Range("nextMtgNo"))
    hasChanged = False

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    hasChanged = True
End Sub

Private Sub Worksheet_Deactivate()

    If hasChanged Then Call UpdateMtgRosterTtblAtDB

End Sub

Sub UpdateMtgRosterTtblAtDB()
        
    ' SQLiteDB 接続処理
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myStmtHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    #End If

    ' Open the database - getting a DbHandle back
    Call SQLite3Open(ThisWorkbook.Path & SQLITE_PATH, myDbHandle)
    
    Dim lstRow As ListRow
    Dim colIndex As Long
        
    'SQLiteDB書込
    Call SQLite3ExecuteNonQuery(myDbHandle, "BEGIN TRANSACTION")
    
    Dim aryRec As Variant
    Dim val As Variant
    Dim strSQL As String
    With Range(ROSTER_TTBL).ListObject
        For Each lstRow In .ListRows
            ReDim aryRec(1 To 5)
            strSQL = "Insert Into tMtgRoster (MTG_No,参加者ID,調整時参加可否,当日出欠予定,当日出欠) Values" & vbLf
            aryRec(1) = Range("rosterMtgNo")
            For colIndex = 1 To .ListColumns.Count
                val = lstRow.Range(, colIndex)
                If val = "" Then val = "null"
                Select Case colIndex
                    Case eRoster.参加者ID
                        aryRec(2) = "'" & val & "'"
                    Case eRoster.調整時参加可否
                        aryRec(3) = val
                    Case eRoster.出欠予定
                        aryRec(4) = val
                    Case eRoster.当日出欠
                        aryRec(5) = val
                End Select
            Next
            strSQL = strSQL & "(" & Join(aryRec, vbLf & ",") & ")"
            strSQL = strSQL & "On Conflict(MTG_No,参加者ID) Do Update Set" & vbLf
            strSQL = strSQL & "調整時参加可否 = " & aryRec(3) & vbLf
            strSQL = strSQL & ",当日出欠予定 = " & aryRec(4) & vbLf
            strSQL = strSQL & ",当日出欠 = " & aryRec(5)
            
            Call SQLite3PrepareV2(myDbHandle, strSQL, myStmtHandle)
            Call SQLite3Step(myStmtHandle)
        Next
    End With
    
    Call SQLite3Finalize(myStmtHandle)
    Call SQLite3ExecuteNonQuery(myDbHandle, "COMMIT TRANSACTION")

    strSQL = "UPDATE mMember" & vbLf
    strSQL = strSQL & "SET isActive = act.isActive" & vbLf
    strSQL = strSQL & "FROM (" & vbLf
    strSQL = strSQL & "    SELECT" & vbLf
    strSQL = strSQL & "        参加者ID" & vbLf
    strSQL = strSQL & "        ,sign(sum(CASE" & vbLf
    strSQL = strSQL & "            WHEN mm.入会年月日 > date('now', '-1 year') THEN 1" & vbLf
    strSQL = strSQL & "            WHEN ifnull(当日出欠,0) > 0 THEN 当日出欠" & vbLf
    strSQL = strSQL & "            WHEN ifnull(当日出欠予定,0) > 0 THEN 当日出欠予定" & vbLf
    strSQL = strSQL & "            WHEN ifnull(調整時参加可否,0) > 0 THEN 調整時参加可否" & vbLf
    strSQL = strSQL & "            ELSE 0" & vbLf
    strSQL = strSQL & "        END)) AS isActive" & vbLf
    strSQL = strSQL & "    FROM tMtgRoster AS tmr" & vbLf
    strSQL = strSQL & "    LEFT JOIN tmMeeting AS tm ON tm.MTG_No = tmr.MTG_No" & vbLf
    strSQL = strSQL & "    LEFT JOIN mMember AS mm ON mm.ID = tmr.参加者ID" & vbLf
    strSQL = strSQL & "    WHERE tm.日付 > date('now', '-1 year')" & vbLf
    strSQL = strSQL & "    Group BY" & vbLf
    strSQL = strSQL & "    参加者ID" & vbLf
    strSQL = strSQL & ") AS act" & vbLf
    strSQL = strSQL & "WHERE mMember.ID = act.参加者ID"
    
    Call SQLite3PrepareV2(myDbHandle, strSQL, myStmtHandle)
    Call SQLite3Step(myStmtHandle)
    Call SQLite3Finalize(myStmtHandle)
    
    Call SQLite3Close(myDbHandle)
    
    hasChanged = False
    
End Sub

Sub DisplayNextMtg()
    
    If hasChanged Then Call UpdateMtgRosterTtblAtDB
    Call CreateOrDisplayNextMTG(wksHome.Range("nextMtgNo"))
    
End Sub

Sub DisplayPreviousMtg()

    If hasChanged Then Call UpdateMtgRosterTtblAtDB
    Call CreateOrDisplayNextMTG(Range("rosterMtgNo") - 1)

End Sub

Sub DisplayLaterMtg()
    
    If hasChanged Then Call UpdateMtgRosterTtblAtDB
    Call CreateOrDisplayNextMTG(Range("rosterMtgNo") + 1)

End Sub
