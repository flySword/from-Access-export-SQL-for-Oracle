Attribute VB_Name = "ģ��1"
Option Compare Database
Function CreateSQLString(ByVal FilePath As String) As Boolean
'���������ݵ�ǰMDB�еı���һ�� *.jetsql �ű�
'������������������Ľ����������Ϊ JET SQL DDL ��䲻֧��һЩ ACCESS ���е����ԣ�DAO֧�֣�
'This function create a "*.jetsql" script based on current mdb tables.
'This function is not the BEST, because the JET SQL DDL never support some DAO property.
    Dim MyTableName As String
    Dim MyFieldName As String
    Dim MyDB As New ADOX.Catalog
    Dim MyTable As ADOX.Table
    Dim MyField As ADOX.Column
    Dim pro
    Dim iC As Long
    Dim strField() As String
    Dim strFieldTm As String
    
    Dim strKey As String
    Dim strSQL As String
    Dim strSQLScript As String
    Dim objFile, stmFile
    Dim strText As String
    Dim dct As Dictionary
    
    Dim cn, rs, i
    Dim strTemp As String
    Dim strErrSQL As String
    Dim isErr As Integer
    
    
    
On Error GoTo CreateSQLScript_Err
    MyDB.ActiveConnection = CurrentProject.Connection
    For Each MyTable In MyDB.Tables
        If MyTable.Type = "TABLE" Then
        'ָ��������ͣ����硰TABLE������SYSTEM TABLE����GLOBAL TEMPORARY�����ߡ�ACCESS TABLE����
        'ADOX �޷��жϸñ��Ƿ��Ѿ���ɾ�����������ַ�ʽ�жϣ�
        '����һ������ DAO��
        'If CurrentDb.TableDefs(strTableName).Attributes = 0 Then
        '�������������ж� ADOX.Table.Type �Ļ��������ж�������
        'If Left(MyTable.Name, 7) <> "~TMPCLP" Then
 
                    
            
            isErr = 0
            strSQL = "create table " & MyTable.Name & "( "
            
            If MyTable.Name = "GEOSPATIALCHARTLABEL" Then
                isErr = 0
            End If
                
            
            
            
            
            Set dct = CreateObject("Scripting.Dictionary")
            Set cn = CurrentProject.Connection
            Set rs = cn.OpenSchema(4, Array(Null, Null, MyTable.Name))
            
            
            With rs
                Do While Not .EOF
                    strTemp = .Fields("COLUMN_NAME").Value
                          
                          
                    '�˴����������������
                    Dim p As String
                    Select Case .Fields("DATA_TYPE").Value
                        Case 11
                            p = " NUMERIC(1)"    'yesno ��Ϊ NUMERIC(1)
                        Case 6
                            p = " money"
                        Case 7
                            p = " DATE"     'datetime ��Ϊ DATE
                        Case 5
                            p = " FLOAT"    'or " FLOAT"
                            'p = " NUMERIC"   'decimal to NUMERIC for oracle
                            'p = p & "(" & .Fields("NUMERIC_PRECISION").Value & "," & .Fields("NUMERIC_SCALE").Value & ")"
                        Case 72
                            'JET SQL DDL ����޷��������Զ���� GUID���ֶΣ�������ʱ��
                            '[d] GUID default GenGUID() ���沿�ֹ��ܣ������뿴����
                            '�����JET SQL DDL�����Զ����GUID�ֶ�
                            'http://access911.net/?kbid;72FABE1E17DCEEF3
                                 p = " GUIDdfdfgd"
                        Case 3
                                  p = " smallint"
                        
                        Case 205
                            p = " BLOB"  'image to BLOB for oracle
                        Case 203
                            p = " varchar(1024)"   'memo to varchar for oracle  'Access "HyperLink" field is also a MEMO data type.
                            'ACCESS �ĳ�������Ҳ�� MEMO ���͵�
                        Case 131
                            p = " NUMERIC"   'decimal to NUMERIC for oracle
                            p = p & "(" & .Fields("NUMERIC_PRECISION").Value & "," & .Fields("NUMERIC_SCALE").Value & ")"
                        Case 4
                            p = " NUMERIC(10,3)"  'single to NUMERIC(10,3) for oracle     'or " REAL"
                        Case 2
                            p = " smallint"
                        Case 17
                            p = " varchar(6)"    ' byte ��Ϊ varchar(6) for oracle
                        Case 202
                            p = " varchar"   ' nvarchar to varchar for oracle
                            If .Fields("CHARACTER_MAXIMUM_LENGTH").Value = 0 Then
                                 p = p & "(1024)"
                             Else
                                 p = p & "(" & .Fields("CHARACTER_MAXIMUM_LENGTH").Value & ")"
                            End If
                            
                        Case 130
                            'ָʾһ���� Null ��ֹ�� Unicode �ַ��� (DBTYPE_WSTR)�� �������������� ACCESS ��������޷���Ƴ����ġ�
                            '20100826 ����
                            p = " varchar"
                             If .Fields("CHARACTER_MAXIMUM_LENGTH").Value = 0 Then
                                 p = p & "(1024)"
                             Else
                                 p = p & "(" & .Fields("CHARACTER_MAXIMUM_LENGTH").Value & ")"
                            End If
                            
                        Case 128
                            p = " BLOB"
                            
                            
                            
                        Case Else
                             isErr = 1
                            p = " (" & .Fields("DATA_TYPE").Value & " ������δ�������ͣ�����"
                    End Select
                  ' p = " " & objField.Name & " " & p
                   'If IsEmpty(objField.Properties("Default")) = False Then
                       'p = p & " default " & objField.Properties("Default")
                   'End If
                   'If objField.Properties("Nullable") = False Then
                       'p = p & " not null"
                   'End If
                   'SQLField = p
                    
                    strTemp = strTemp & " " & p
                    dct.Add .Fields("ORDINAL_POSITION").Value, strTemp
                    
                   ' strFieldTm = SQLField(.Field
                        
                    'strSQL = strSQL & strTemp
                    
                    
            '                        strField(iC) = SQLField(MyField) ' �ļ�����
            '                        iC = iC + 1
               
                    .MoveNext
                      
                Loop
                .Close
            End With
                
          '          MsgBox strSQL
            
                    
            For i = 1 To dct.Count
               strSQL = strSQL & " " & dct(i) & " , "
            Next
            
            strSQL = Left(strSQL, Len(strSQL) - 2)  ' ȥĩβ�Ķ���
            strKey = SQLKey(MyTable)
            If Len(strKey) <> 0 Then
                strSQL = strSQL & "," & strKey
            End If
            strSQL = strSQL & " );" & vbCrLf
            
            Set dct = Nothing
            Set rs = Nothing
            cn.Close
            Set cn = Nothing
                    
               ' End If
                
                
                
    '            For Each MyField In MyTable.Columns
    '                ReDim Preserve strField(iC)
    '                strField(iC) = SQLField(MyField) ' �ļ�����
    '                iC = iC + 1
    '            Next
    '            strSQL = strSQL & Join(strField, ",")
    '            '��ȡ��ǰ����ֶ���Ϣ���������³�ʼ�� strField ����
    '            iC = 0
    '            ReDim strField(iC)
    '            '�������Ϣ
    '            strKey = SQLKey(MyTable)
    '            If Len(strKey) <> 0 Then
    '                strSQL = strSQL & "," & strKey
    '            End If
    '            strSQL = strSQL & ");" & vbCrLf
     '           strSQLScript = strSQLScript & strSQL
     '           'Debug.Print SQLIndex(MyTable)      'Never support the INDEX,to be continued...
      '          '��δ֧�� index �ű���δ�����...
    
            
            If isErr = 0 Then
                 strSQLScript = strSQLScript & strSQL
            Else
                strErrSQL = strErrSQL & strSQL
            End If
        End If
    Next
    
   
    
    Set MyDB = Nothing
    'create the Jet SQL Script File
    Set objFile = CreateObject("Scripting.FileSystemObject")
    Set stmFile = objFile.CreateTextFile(FilePath, True)
    stmFile.Write strSQLScript

    stmFile.Write vbCrLf & vbCrLf & vbCrLf
    If Len(strErrSQL) <> 0 Then
        stmFile.Write "access�������ʹ��룬��ҪΪOLE��������" & vbCrLf
    End If
    
    stmFile.Write strErrSQL
    stmFile.Close
    Set stmFile = Nothing
    Set objFile = Nothing
    CreateSQLScript = True
CreateSQLScript_Exit:
    Exit Function
CreateSQLScript_Err:
    MsgBox Err.Description, vbExclamation
    CreateSQLScript = False
    Resume CreateSQLScript_Exit
End Function





Function SQLKey(ByVal objTable As ADOX.Table)
'���� ADOX �����йء������� JET SQL DDL �Ӿ�
'Reference ADOX and create the JET SQL DDL clause about the "Key"
    Dim MyKey As ADOX.Key
    Dim MyKeyColumn As ADOX.Column
    Dim strKey As String
    Dim strColumns() As String
    Dim strKeys() As String
    Dim i As Long
    Dim iC As Long
    For Each MyKey In objTable.Keys
    
        Select Case MyKey.Type
        Case adKeyPrimary
            strKey = "Primary KEY "
        Case adKeyForeign
            Exit For
            
            'strKey = "FOREIGN KEY "
        Case adKeyUnique
            Exit For
            
            'strKey = "UNIQUE "
        End Select
        For Each MyKeyColumn In MyKey.Columns
            ReDim Preserve strColumns(iC)
            strColumns(iC) = " " & MyKeyColumn.Name & " "
            iC = iC + 1
        Next
        ReDim Preserve strKeys(i)
        strKeys(i) = strKey & "(" & Join(strColumns, ",") & ")"
        '��ȡ��Ϣ��������ʼ������
        iC = 0
        ReDim strColumns(iC)
        i = i + 1
        
    Next
    SQLKey = Join(strKeys, ",")
End Function


Function SQLField(ByVal objField As Fields)
'���� ADOX �����йء��ֶΡ��� JET SQL DDL �Ӿ�
'Reference ADOX and create the JET SQL DDL clause about the "Field"
    Dim p As String
    Select Case objField.Type
        Case 11
            p = " NUMERIC(1)"    'yesno ��Ϊ NUMERIC(1)
        Case 6
            p = " money"
        Case 7
            p = " DATE"     'datetime ��Ϊ DATE
        Case 5
            p = " FLOAT"    'or " Double"
        Case 72
            'JET SQL DDL ����޷��������Զ���� GUID���ֶΣ�������ʱ��
            '[d] GUID default GenGUID() ���沿�ֹ��ܣ������뿴����
            '�����JET SQL DDL�����Զ����GUID�ֶ�
            'http://access911.net/?kbid;72FABE1E17DCEEF3
            If objField.Properties("Autoincrement") = True Then
                p = " autoincrement GUID"
            Else
                p = " GUID"
            End If
        Case 3
            If objField.Properties("Autoincrement") = False Then
                p = " smallint"
            Else
                p = " AUTOINCREMENT(1," & objField.Properties("Increment") & ")"
            End If
        Case 205
            p = " BLOB"  'image to BLOB for oracle
        Case 203
            p = " varchar(1024)"   'memo to varchar for oracle  'Access "HyperLink" field is also a MEMO data type.
            'ACCESS �ĳ�������Ҳ�� MEMO ���͵�
        Case 131
            p = " NUMERIC"   'decimal to NUMERIC for oracle
            p = p & "(" & objField.Precision & "," & objField.NumericScale & ")"
        Case 4
            p = " NUMERIC(10,3)"  'single to NUMERIC(10,3) for oracle     'or " REAL"
        Case 2
            p = " smallint"
        Case 17
            p = " varchar(6)"    ' byte ��Ϊ varchar(6) for oracle
        Case 202
            p = " varchar"   ' nvarchar to varchar for oracle
            p = p & "(" & objField.DefinedSize & ")"
        Case 130
            'ָʾһ���� Null ��ֹ�� Unicode �ַ��� (DBTYPE_WSTR)�� �������������� ACCESS ��������޷���Ƴ����ġ�
            '20100826 ����
            p = " char"
            p = p & "(" & objField.DefinedSize & ")"
        Case Else
            p = " (" & objField.Type & " Unknown,You can find it in ADOX's help. Please Check it.)"
    End Select
    p = " " & objField.Name & " " & p
    If IsEmpty(objField.Properties("Default")) = False Then
        p = p & " default " & objField.Properties("Default")
    End If
    If objField.Properties("Nullable") = False Then
        p = p & " not null"
    End If
    SQLField = p
End Function



Public Function GetfilType(Tbname As String, Filname As String) As String
'����:����ֶ�����
Dim rs As New ADODB.Recordset
rs.Open Tbname, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
GetfilType = rs(Filname).Type
End Function





Sub RunTest_CreateScript()
    CreateSQLString "E:\desktop\temp\���ʽ���\oracle_sql.sql"
    
End Sub





Sub test()
    Dim GetfilType As String
    
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim cn As ADODB.Connection
    Dim FN As ADODB.Field
    
    Set cn = CurrentProject.Connection
    
    Dim i As Long
    Set rs = cn.OpenSchema(20)
    ' rs.Open "DB03MineSvy_01TC_BaseInfo", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    
    rs.MoveNext
    If UCase(rs("TABLE_TYPE")) = "TABLE" Then
        MsgBox "��:" & rs("TABLE_NAME")
    End If
    
    rs1.Open "DB03MineSvy_01TC_BaseInfo", cn
    For Each FN In rs.Fields
        MsgBox FN.Name
    Next
    
    
    
  
    'MsgBox GetfilType


End Sub




Sub test2()
    Dim adSchemaColumns As Integer
    adSchemaColumns = 4
    Dim cn, rs, dct, i
    
    Set dct = CreateObject("Scripting.Dictionary")
    Set cn = CurrentProject.Connection
    Set rs = cn.OpenSchema(4, Array(Null, Null, "DB03MineSvy_01TC_BaseInfo"))
    

    Dim aa As Integer
    

    
    
    MsgBox rs.Fields(11).Name
    MsgBox rs.Fields(2).Value
    
    
    aa = 0
    With rs
        Do While Not .EOF
            dct.Add .Fields("ORDINAL_POSITION").Value, .Fields("COLUMN_NAME").Value & .Fields("DATA_TYPE").Value
            
            'MsgBox .Fields("DATA_TYPE").Value
            
            
            .MoveNext
            
              
        Loop
        .Close
    End With
    
    
    
    

    

    
    For i = 1 To dct.Count
        
       MsgBox dct(i)
    Next
    
    Set dct = Nothing
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Sub
