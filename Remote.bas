Attribute VB_Name = "Remote_Globali"
'
Private Type M_Rem
   Pressed As Boolean
   KeyPress As Integer
   Left As Integer
   Right As Integer
   Up As Integer
   Down As Integer
   OK As Integer
   Escape As Integer
   Zoom As Integer
End Type
Public Remote As M_Rem

Public Function Remote_Set()
  Remote.Right = 38 '39
  Remote.Left = 40 '37
  Remote.OK = 13
  Remote.Escape = 18 '27
End Function
