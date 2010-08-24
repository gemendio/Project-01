Attribute VB_Name = "Global"
Public Function db() As ADODB.Connection
    Dim c As New ADODB.Connection
    c.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" _
                         & "SERVER=localhost;" _
                         & "DATABASE=room_scheduler;" _
                         & "UID=root;" _
                         & "PWD=p@ssword;" _
                         & "OPTION=3"
    Set db = c
End Function

Public Function insertSql(table_name As String, data() As Variant, fields() As String)
    If (UBound(data) = UBound(fields)) And table_name <> "" Then
        insertSql = "INSERT INTO " & table_name _
            & " (" & Join(fields, ", ") & ") " _
            & "VALUES('" & Join(data, "', '") & "')"
    End If
End Function

Public Function deleteSql(table_name As String, where As String)
    If table_name <> "" And where <> "" Then
        deleteSql = "DELETE FROM " & table_name & " " & where
    End If
End Function
'this function is not yet done
Public Function updateSql(table_name As String, data() As Variant, fields() As String, where As String)
    If (UBound(data) = UBound(fields)) And table_name <> "" Then

        Dim field_ubound, i As Integer

        field_ubound = UBound(fields)
        
        Dim set_sql(0 To 6) As String


        For i = 0 To field_ubound
            set_sql(i) = fields(i) & " = '" & data(i)
        Next i


MsgBox Join(set_sql, "', ")
        updateSql = "UPDATE " & table_name _
            & " SET " & Join(sql_set, "', ") & where
    End If
End Function

