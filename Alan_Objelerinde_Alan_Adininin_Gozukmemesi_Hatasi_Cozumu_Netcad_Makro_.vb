Sub Main
Dim i,ada,o
 with Netcad
 for i = 0 to .numobject-1 ' projedeki tum objeleri sirayla tara
 .BackMessage
 .SetMessage i
 set ada = .getobject(i) ' i. objeyi al
 if ada.tag = opline then
 .DrawObject ada, blue
 ada.flags=17
 .putobject i,ada
 end if
next

 .netcadcommand("REGENSINGLE " & .GetCurrentWindow.cll.y & "," &.GetCurrentWindow.cll.x & " " & .GetCurrentWindow.cur.y & "," & .GetCurrentWindow.cur.x)
set ada=nothing
.BackMessage 
msgbox " Alan Adı Gösterimi Tamamlandı [Harita Akademi, Şaban GÜL]",64 ," Tebrikler [Harita Akademi, Şaban GÜL]"
 end with
End Sub
