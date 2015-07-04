Attribute VB_Name = "createInsert"
'/**
' * Inser���̃t�@�C�����쐬����
' */
Sub create_Insert()

    ' SQL������
    Dim ddl As String
    ' �e�[�u����
    Dim tableName As String
    'defSql
    Dim defSql As String
    ' �ŏI�s
    Dim lastRow As Integer
    ' �ŏI��
    Dim lastCol As Integer
    
    '�}������l
    Dim strValue() As String
    
    Dim tempSql As String
    
    '�J�E���^
    Dim i, j As Integer

    ' �t�@�C����
    Dim flName As String
    ' �t�@�C��
    Dim fn As Integer
     
    '�t�@�C�������擾����
    flName = Application.GetSaveAsFilename("", "")

    ' �ŏI�s���擾����
    lastRow = Selection.Rows.Count
    ' �ŏI����擾����
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

    '/* ���p�\�ȃt�@�C���m���𓾂�
    fn = FreeFile()
    '/* �t�@�C���̃I�[�v��
    Open flName For Output As #fn
    
    '/* �ۑ��icrlf�t�j
    Print #fn, ddl
  
    '/* �t�@�C�������
    Close #fn
    
    
End Sub

'/**
' * CSV�̃t�@�C�����쐬����
' */
Sub create_ProperCsv()

    ' SQL������
    Dim csv As String
    
    ' �ŏI�s
    Dim lastRow As Integer
    ' �ŏI��
    Dim lastCol As Integer
    
    '�}������l
    Dim strValue() As String
    
    Dim rowCsv As String
    
    '�J�E���^
    Dim i, j As Integer

    ' �t�@�C����
    Dim flName As String
    ' �t�@�C��
    Dim fn As Integer
     
    '�t�@�C�������擾����
    flName = Application.GetSaveAsFilename("", "")

    ' �ŏI�s���擾����
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    lastRow = Selection.Rows.Count
    ' �ŏI����擾����
    lastCol = Selection.Columns.Count
    
    ReDim strValue(0 To lastCol - 1)
    
    For i = 0 To lastRow - 1
        For j = 0 To lastCol - 1
            
            strValue(j) = ActiveCell.Offset(i, j).Value
        
        Next j
                
        rowCsv = """" & Join(strValue, """,""") & """" & vbCrLf
        csv = csv & rowCsv
    Next i

    '/* ���p�\�ȃt�@�C���m���𓾂�
    fn = FreeFile()
    '/* �t�@�C���̃I�[�v��
    Open flName For Output As #fn
    
    '/* �ۑ��icrlf�t�j
    Print #fn, csv
  
    '/* �t�@�C�������
    Close #fn
    
    
End Sub

