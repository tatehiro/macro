Attribute VB_Name = "createInsert"
'/**
' * Inser文のファイルを作成する
' */
Sub create_Insert()

    ' SQL文字列
    Dim ddl As String
    ' テーブル名
    Dim tableName As String
    'defSql
    Dim defSql As String
    ' 最終行
    Dim lastRow As Integer
    ' 最終列
    Dim lastCol As Integer
    
    '挿入する値
    Dim strValue() As String
    
    Dim tempSql As String
    
    'カウンタ
    Dim i, j As Integer

    ' ファイル名
    Dim flName As String
    ' ファイル
    Dim fn As Integer
     
    'ファイル名を取得する
    flName = Application.GetSaveAsFilename("", "")

    ' 最終行を取得する
    lastRow = Selection.Rows.Count
    ' 最終列を取得する
    lastCol = Selection.Columns.Count
    
    ReDim strValue(0 To lastCol - 1)
    
    tableName = ActiveCell.Offset(0, 0).Value
    ActiveCell.Offset(1, 0).Select
    
    For i = 0 To lastRow - 2
        tempSql = ""
        defSqls = ""
        For j = 0 To lastCol - 1
            
            If i = 0 Then
                If j = 0 Then
                    defSql = "insert into " & tableName & " ("
                End If
                
                defSql = defSql & ActiveCell.Offset(i, j).Value
            
                If j = lastCol - 1 Then
                    defSql = defSql & ") values ("
                Else
                    defSql = defSql & ", "
                End If
                
            Else
                
                If ActiveCell.Offset(i, j).Value <> "" Then
                    strValue(j) = ActiveCell.Offset(i, j).Value
                Else
                    strValue(j) = "NULL"
                End If
                                
            End If
        
        Next j
        
        If i <> 0 Then
            
            tempSql = "'" & Join(strValue, "', '") & "');"
            tempSql = Replace(tempSql, "'NULL'", "NULL")
            tempSql = defSql & tempSql & vbLf
            ddl = ddl & tempSql
        
        End If
    
    Next i

    '/* 利用可能なファイルＮｏを得る
    fn = FreeFile()
    '/* ファイルのオープン
    Open flName For Output As #fn
    
    '/* 保存（crlf付）
    Print #fn, ddl
  
    '/* ファイルを閉じる
    Close #fn
    
    
End Sub

'/**
' * CSVのファイルを作成する
' */
Sub create_ProperCsv()

    ' SQL文字列
    Dim csv As String
    
    ' 最終行
    Dim lastRow As Integer
    ' 最終列
    Dim lastCol As Integer
    
    '挿入する値
    Dim strValue() As String
    
    Dim rowCsv As String
    
    'カウンタ
    Dim i, j As Integer

    ' ファイル名
    Dim flName As String
    ' ファイル
    Dim fn As Integer
     
    'ファイル名を取得する
    flName = Application.GetSaveAsFilename("", "")

    ' 最終行を取得する
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    lastRow = Selection.Rows.Count
    ' 最終列を取得する
    lastCol = Selection.Columns.Count
    
    ReDim strValue(0 To lastCol - 1)
    
    For i = 0 To lastRow - 1
        For j = 0 To lastCol - 1
            
            strValue(j) = ActiveCell.Offset(i, j).Value
        
        Next j
                
        rowCsv = """" & Join(strValue, """,""") & """" & vbCrLf
        csv = csv & rowCsv
    Next i

    '/* 利用可能なファイルＮｏを得る
    fn = FreeFile()
    '/* ファイルのオープン
    Open flName For Output As #fn
    
    '/* 保存（crlf付）
    Print #fn, csv
  
    '/* ファイルを閉じる
    Close #fn
    
    
End Sub

