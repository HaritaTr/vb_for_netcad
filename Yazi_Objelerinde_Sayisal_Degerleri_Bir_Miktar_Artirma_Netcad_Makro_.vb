Sub Main

Dim i
 dim obj
 dim regpoly
 dim bd
 dim sagul,sagul2
 dim elifyaren
 
on error resume next
with Netcad

set BD = Netcad.NewBDialog("Yazı Objelerinde Sayısal Değerleri Bir Miktar Artırma [Harita Akademi, Şaban GÜL]")

BD.GetCombo "tabaka", "İşlem Görecek Tabaka Seçiniz: ","0" , 0
 for i =  1 to .numlayers-1
 BD.AddCombo .LayerNameOf(i)
 next

BD.Getinteger "item","Artış Miktarı Giriniz:",0

BD.GetCheck "tabaka2", "Tüm Tabakalarda İşlem Yapılsın",0
'BD.GetCheck "tabaka3", "Tabaka Yerine Ekrandan Seçilsin",0





 if BD.showmodal then
 sagul=BD.ValueByName("tabaka")
sagul2= BD.ValueByName("tabaka2")
 else
 exit sub
 end if



if sagul2=1 then
 .SetFilter nothing, ARRAY(), ARRAY(otext)
 else
 .SetFilter nothing, ARRAY(sagul), ARRAY(otext)
 end if

elifyaren=0

DO
 SET OBJ=.GETNEXTOBJECT

IF OBJ IS NOTHING THEN
 EXIT DO
 ELSE
 END IF

obj.s=obj.s+BD.ValueByName("item")
elifyaren=elifyaren+1
 .PUTOBJECT .CUROBJPOS,OBJ

LOOP


.netcadcommand("REGEN")
msgbox "İşlem Başarıyla Tamamlandı." & elifyaren & " adet yazı değiştirildi" , 64 , "[Harita Akademi, Şaban GÜL]"
end with
End Sub
