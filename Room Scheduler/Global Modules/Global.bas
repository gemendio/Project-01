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
        Dim sql_set() As String
        Dim field As Variant
        Dim n As Integer
        
        n = 0
        
        For Each field In fields
            sql_set(n) = field & " = " & data(n)
            n = n + 1
        Next

        updateSql = "UPDATE " & table_name _
            & " SET " & Join(sql_set, ", ") & where
        
End Function

