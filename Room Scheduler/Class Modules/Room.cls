VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModelRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_id As Integer
Private m_name As String

Public Property Let ID(value As Integer)
    m_id = value
End Property

Public Property Get ID() As Integer
    ID = m_id
End Property

Public Property Let Name(value As String)
    m_name = value
End Property

Public Property Get Name() As String
    Name = m_name
End Property

Public Sub Upsert()
    Dim conn As New ADODB.Connection
    conn = db
    
    conn.Open
    
    If ID = 0 Then
        conn.Execute "INSERT INTO rooms (name) VALUES ('" & Name & "')"
    Else
        conn.Execute "UPDATE rooms SET name = '" & Name & "' WHERE id = " & ID
    End If
    
    conn.Close
End Sub



