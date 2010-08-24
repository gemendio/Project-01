Attribute VB_Name = "Global"
Public Function db() As ADODB.Connection
    Dim c As New ADODB.Connection
    c.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" _
                         & "SERVER=localhost;" _
                         & "DATABASE=room_scheduler;" _
                         & "UID=root;" _
                         & "PWD=lordofwar;" _
                         & "OPTION=3"
    Set db = c
End Function


