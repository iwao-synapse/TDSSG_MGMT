Attribute VB_Name = "fGetRecordsFromDB_v1_01"
Option Explicit

Function GetRecordsFromDB(strSQL As String, Optional colIndexFrom As Long, Optional titleRow As Boolean) As Variant

    Dim records As Variant
    records = Array()   'Array(Array()) Typeでデータ格納

    'Default設定：Columnは、0IndexBaseでレコード（配列）格納
    'Option設定：Columnは、option入力された数値を起点にレコード（配列）格納

    ' SQLiteDB 接続処理
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myStmtHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    #End If
    Dim retVal As Long
            
    ' Open the database - getting a DbHandle back
    Call SQLite3Open(ThisWorkbook.Path & SQLITE_PATH, myDbHandle)

    ' getting a StmtHandle back
    Call SQLite3PrepareV2(myDbHandle, strSQL, myStmtHandle)
    retVal = SQLite3Step(myStmtHandle)
    If retVal = SQLITE_DONE Or retVal = SQLITE_MISUSE Then
        GetRecordsFromDB = Empty
        Exit Function
    End If

    Dim colMax As Long
    Dim col As Long
    colMax = SQLite3ColumnCount(myStmtHandle)
    
    Dim record As Variant
    ReDim record(colIndexFrom To colIndexFrom + colMax - 1)
        
    ' タイトル列処理
    If titleRow Then
        For col = 0 To colMax - 1
            record(col + colIndexFrom) = SQLite3ColumnName(myStmtHandle, col)
        Next col
        ReDim records(UBound(records) + 1)
        records(UBound(records)) = record
    End If
    
    Do Until retVal = SQLITE_DONE Or retVal = SQLITE_MISUSE
        
        Dim colType As Long
        Dim colValue As Variant
            
        '1列ずつ配列格納
        For col = 0 To colMax - 1
            colType = SQLite3ColumnType(myStmtHandle, col)
            colValue = ColumnValue(myStmtHandle, col, colType)
            record(col + colIndexFrom) = colValue
        Next col
        
        DoEvents
        
        ' 1レコードずつ配列格納  ← Array(Array()) Type
        ReDim Preserve records(UBound(records) + 1)
        records(UBound(records)) = record
        
        ' Move to next row
        retVal = SQLite3Step(myStmtHandle)
    
    Loop
    
    ' Finalize (delete) the statement
    Call SQLite3Finalize(myStmtHandle)
    ' Close the database
    Call SQLite3Close(myDbHandle)
    
    GetRecordsFromDB = records
    
End Function

