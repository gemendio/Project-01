VERSION 5.00
Begin VB.Form Rooms 
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Rooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim room As New ModelRoom
    Dim user As New ModelUser
    Dim section As New ModelSection
    Dim o As Object
    
    
Set o = getAll("users")
MsgBox o
    'MsgBox user.FirstName & " " & user.MiddleName & " " & user.LastName
    
    
End Sub
