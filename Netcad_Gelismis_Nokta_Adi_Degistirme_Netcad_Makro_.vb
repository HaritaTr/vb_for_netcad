Sub Main

Dim i
dim obj
dim regpoly
dim bd ,bd2
dim sagul,sagul2
dim sabangul1,sabangul2,sabangul3,sabangul4,sabangul5
with Netcad

set BD2 = Netcad.NewBDialog("Gelişmiş Nokta İsmi Değiştirme [Harita Akademi, Şaban GÜL]")
BD2.PutPrompt "Yapılan İşlemde Öncelikle Nokta İsmine Sayısal Değer Ekleme"
BD2.PutPrompt "Ardından Önce Başına, Sonra Sonuna Değer Ekleme ve en son "
BD2.PutPrompt "Bul Değiştir işlemleri yapılır. Metin Ekleme en fazla 3 karakter sınırındadır"
BD2.PutPrompt " "
BD2.PutPrompt "NETCADDE NOKTA SAYISI EN FAZLA 10 KARAKTERDİR."
BD2.PutPrompt " "

BD2.GetCombo "tabaka", "İşlem Görecek Tabaka Seçiniz: ", "0", 0
for i = 1 to .numlayers-1
BD2.AddCombo .LayerNameOf(i)
next

BD2.GetCheck "tabaka2", "Tüm Tabakalarda İşlem Yapılsın",0

BD2.PutPrompt " "
BD2.PutPrompt "Devam Etmek İçin Tamam Butonuna Basınız."
if BD2.showmodal then

else
exit sub
end if



set BD = Netcad.NewBDialog("Gelişmiş Nokta İsmi Değiştirme [Harita Akademi, Şaban GÜL]")

BD.PutPrompt "Bir Noktanın Başına ve Sonuna Metin Ekleme"
BD.Getstring "elifyaren1","Başına Ekle:","", 3
BD.Getstring "elifyaren2","Sonuna Ekle:","", 3
BD.PutPrompt " "
BD.PutPrompt "Bir Noktanın Adında Bul Değiştir Yapma"
BD.Getstring "elifyaren3","Eski Metin:","", 50
BD.Getstring "elifyaren4","Yeni Metin:","", 50
BD.PutPrompt " "
BD.Getinteger "item","Nokta İsmine Sayısal Değer Ekle/Çıkart:",0



sagul=BD2.ValueByName("tabaka")
sagul2= BD2.ValueByName("tabaka2")
if BD.showmodal then

sabangul1= BD.ValueByName("elifyaren1")
sabangul2= BD.ValueByName("elifyaren2")
sabangul3= BD.ValueByName("elifyaren3")
sabangul4= BD.ValueByName("elifyaren4")
sabangul5= BD.ValueByName("item")
if sagul2=1 then
.SetFilter nothing, ARRAY(), ARRAY(opoint)
else
.SetFilter nothing, ARRAY(.FOUNDLAYER(sagul)), ARRAY(opoint)
end if

on error resume next
DO
SET OBJ=.GETNEXTOBJECT

IF OBJ IS NOTHING THEN
EXIT DO
ELSE
END IF

if sabangul5=0 then
else
obj.pname=obj.pname+sabangul5
end if

obj.pname= sabangul1 & obj.pname
obj.pname=obj.pname & sabangul2

obj.pname=replace(obj.pname,sabangul3,sabangul4)



.PUTOBJECT .CUROBJPOS,OBJ



LOOP



end if

.netcadcommand("REGEN")

end with

End Sub
