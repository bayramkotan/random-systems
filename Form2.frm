VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entropi-Hal fonksiyonu"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
  Dim i As Integer
  Form2.Cls
  
  Me.Scale (0, ent_dizi(hal - 1) * 1.05)-(hal * 1.05, 0)
  Line (0, ent_dizi(hal - 1) * 1.009)-(hal * 1.05, ent_dizi(hal - 1) * 1.009), vbRed
  For i = 0 To hal - 2
   Line (i + 1, ent_dizi(i))-(i + 2, ent_dizi(i + 1))
  Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CInt(X) > 300 Then Exit Sub
If ent_dizi(CInt(X)) <> 0 Then
Me.Caption = "x=" & X & " y=" & Y & "            " & CInt(X) & ". hal için Entropi=" & ent_dizi(CInt(X))
Else
Me.Caption = "x=" & X & " y=" & Y
End If
End Sub

Private Sub Form_Resize()
Me.Cls
Form_Activate
End Sub
