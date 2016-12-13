Attribute VB_Name = "模块1"
Option Compare Database

' microsoft scripting running
' microsoft ActiveX data objects 2.8 library
' microsoft ADO Ext. 6.0 for DLL and security


Function CreateSQLString(ByVal FilePath As String) As Boolean
'本函数根据当前MDB中的表创建一个 *.jetsql 脚本
'这个函数不是最完美的解决方案，因为 JET SQL DDL 语句不支持一些 ACCESS 特有的属性（DAO支持）
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
        '指定表的类型，例如“TABLE”、“SYSTEM TABLE”或“GLOBAL TEMPORARY”或者“ACCESS TABLE”。
        'ADOX 无法判断该表是否已经被删除，还有两种方式判断，
        '方法一：（用 DAO）
        'If CurrentDb.TableDefs(strTableName).Attributes = 0 Then
        '方法二：（在判断 ADOX.Table.Type 的基础上再判定表名）
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
                          
                          
                    '此处处理各种数据类型
                    Dim p As String
                    Select Case .Fields("DATA_TYPE").Value
                        Case 11
                            p = " NUMERIC(1)"    'yesno 改为 NUMERIC(1)
                        Case 6
                            p = " money"
                        Case 7
                            p = " DATE"     'datetime 改为 DATE
                        Case 5
                            p = " FLOAT"    'or " FLOAT"
                            'p = " NUMERIC"   'decimal to NUMERIC for oracle
                            'p = p & "(" & .Fields("NUMERIC_PRECISION").Value & "," & .Fields("NUMERIC_SCALE").Value & ")"
                        Case 72
                            'JET SQL DDL 语句无法创建“自动编号 GUID”字段，这里暂时用
                            '[d] GUID default GenGUID() 代替部分功能，详情请看文章
                            '如何用JET SQL DDL创建自动编号GUID字段
                            'http://access911.net/?kbid;72FABE1E17DCEEF3
                                 p = " GUIDdfdfgd"
                        Case 3
                                  p = " smallint"
                        
                        Case 205
                            p = " BLOB"  'image to BLOB for oracle
                        Case 203
                            p = " varchar(1024)"   'memo to varchar for oracle  'Access "HyperLink" field is also a MEMO data type.
                            'ACCESS 的超级链接也是 MEMO 类型的
                        Case 131
                            p = " NUMERIC"   'decimal to NUMERIC for oracle
                            p = p & "(" & .Fields("NUMERIC_PRECISION").Value & "," & .Fields("NUMERIC_SCALE").Value & ")"
                        Case 4
                            p = " NUMERIC(10,3)"  'single to NUMERIC(10,3) for oracle     'or " REAL"
                        Case 2
                            p = " smallint"
                        Case 17
                            p = " varchar(6)"    ' byte 改为 varchar(6) for oracle
                        Case 202
                            p = " varchar"   ' nvarchar to varchar for oracle
                            If .Fields("CHARACTER_MAXIMUM_LENGTH").Value = 0 Then
                                 p = p & "(1024)"
                             Else
                                 p = p & "(" & .Fields("CHARACTER_MAXIMUM_LENGTH").Value & ")"
                            End If
                            
                        Case 130
                            '指示一个以 Null 终止的 Unicode 字符串 (DBTYPE_WSTR)。 这种数据类型用 ACCESS 设计器是无法设计出来的。
                            '20100826 新增
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
                            p = " (" & .Fields("DATA_TYPE").Value & " ！！！未处理类型！！！"
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
                    
                    
            '                        strField(iC) = SQLField(MyField) ' 文件类型
            '                        iC = iC + 1
               
                    .MoveNext
                      
                Loop
                .Close
            End With
                
          '          MsgBox strSQL
            
                    
            For i = 1 To dct.Count
               strSQL = strSQL & " " & dct(i) & " , "
            Next
            
            strSQL = Left(strSQL, Len(strSQL) - 2)  ' 去末尾的逗号
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
    '                strField(iC) = SQLField(MyField) ' 文件类型
    '                iC = iC + 1
    '            Next
    '            strSQL = strSQL & Join(strField, ",")
    '            '获取当前表的字段信息后立即重新初始化 strField 数组
    '            iC = 0
    '            ReDim strField(iC)
    '            '加入键信息
    '            strKey = SQLKey(MyTable)
    '            If Len(strKey) <> 0 Then
    '                strSQL = strSQL & "," & strKey
    '            End If
    '            strSQL = strSQL & ");" & vbCrLf
     '           strSQLScript = strSQLScript & strSQL
     '           'Debug.Print SQLIndex(MyTable)      'Never support the INDEX,to be continued...
      '          '暂未支持 index 脚本，未完待续...
    
            
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
        stmFile.Write "access错误类型代码，主要为OLE对象类型" & vbCrLf
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
'调用 ADOX 生成有关“键”的 JET SQL DDL 子句
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
        '获取信息后，立即初始化数组
        iC = 0
        ReDim strColumns(iC)
        i = i + 1
        
    Next
    SQLKey = Join(strKeys, ",")
End Function




Sub RunTest_CreateScript()
    CreateSQLString "E:\desktop\temp\地质建库\oracle_sql.sql"
    
End Sub


