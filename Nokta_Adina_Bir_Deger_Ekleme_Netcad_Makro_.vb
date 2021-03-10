
Sub Main

Dim i
 dim obj
 dim regpoly
 dim bd
 dim sagul,sagul2

with Netcad

set BD = Netcad.NewBDialog("Nokta İsimlerini Bir Miktar Artırma [Harita Akademi, Şaban GÜL]")

BD.Getinteger "item","Artış Miktarı Giriniz:",0
 BD.GetCombo "tabaka", "İşlem Görecek Tabaka Seçiniz: ", "0", 0
 for i = 1 to .numlayers-1
 BD.AddCombo .LayerNameOf(i)
 next

BD.GetCheck "tabaka2", "Tüm Tabakalarda İşlem Yapılsın",0



sagul=BD.ValueByName("tabaka")
 sagul2= BD.ValueByName("tabaka2")
 if BD.showmodal then
 if sagul2=1 then
 .SetFilter nothing, ARRAY(), ARRAY(opoint)
 else
 .SetFilter nothing, ARRAY(.FOUNDLAYER(sagul)), ARRAY(opoint)
 end if



DO
 SET OBJ=.GETNEXTOBJECT

IF OBJ IS NOTHING THEN
 EXIT DO
 ELSE
 END IF

obj.pname=obj.pname+BD.ValueByName("item")
 .PUTOBJECT .CUROBJPOS,OBJ



LOOP



end if

.netcadcommand("REGEN")

end with

End Sub
