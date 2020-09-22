Attribute VB_Name = "Globali_Image"
 Public Const pi = 3.141592
 Public ClrTrasp As Double
'

 Public Const Img_Tot As Long = 9
 Public Const Img_Tot_Sub As Long = 4
 Public Const ImgQnt As Long = 9   ' da 0 a 9 = 10 IMAGINI
 Public Const Img_Step As Single = 3
 Public Rotazione(Img_Tot) As Long
 Public Img(Img_Tot, 40) As Picture
 Public ImgTit(Img_Tot) As Picture
 Public ImgFondo(Img_Tot) As Picture
 Public Img_Sub(Img_Tot, Img_Tot_Sub) As Picture

 Public TxtSub(Img_Tot, Img_Tot_Sub) As String
'
Private Type M_Img
   TrspCol(Img_Tot) As Long
   Totale(Img_Tot) As Long          ' Da 0 a 20
   Corrente As Long
   Center_X As Long
   Center_Y As Long
   Back_X As Long
   Back_Y As Long
End Type
Public Imagine As M_Img
'
Private Type M_Titolo
   TrspCol As Long
End Type
Public Titolo As M_Titolo
'
Private Type M_Menu
   Ready As Boolean
   select As Long
   Submenu(Img_Tot) As Boolean
End Type
Public Menu As M_Menu

'
Function CaricaIcone(Num As Long)
 Dim Nfile As String
 '
 For i = 0 To 40
   On Error GoTo Err:
   Nfile = App.Path + "\Bmp" + "\Bmp " + Trim(Str(Num)) + "-" + Trim(Str(i)) + ".Bmp"
   Set Img(Num, i) = LoadPicture(Nfile)
   Imagine.Totale(Num) = i
   Next i
   Exit Function
Err:
   On Error GoTo 0
End Function
'
Function CaricaIconeSub(Num As Long, NumSub As Long)
 Dim Nfile As String
 On Error Resume Next
  Nfile = App.Path & "\Bmp" & "\BmpSub " & Trim(Str(Num)) & "-" & Trim(Str(NumSub)) & ".Bmp"
   Set Img_Sub(Num, NumSub) = LoadPicture(Nfile)
On Error GoTo 0
End Function
'
Function CaricaTitoli(Num As Long)
 Dim Nfile As String
 On Error Resume Next
 '
   Nfile = App.Path + "\Bmp" + "\Titolo (" + Trim(Str(Num)) + ").Bmp"
   Set ImgTit(Num) = LoadPicture(Nfile)
On Error GoTo 0
End Function
Function CaricaFondi(Num As Long)
Dim Nfile As String
' On Error Resume Next
 '
   Nfile = App.Path + "\Bmp" + "\Fondo (" + Trim(Str(Num)) + ").Bmp"
   Set ImgFondo(Num) = LoadPicture(Nfile)
On Error GoTo 0
End Function
'
Function CaricaTestiSub()
   TxtSub(1, 0) = "PROGDVB IRDETO"
   TxtSub(1, 1) = "PROGDVB S2 EMU"
   TxtSub(1, 2) = "MYTHEATRE"
   TxtSub(1, 3) = "DVBDREAM"
   TxtSub(1, 4) = "ALT-DVB"
End Function
