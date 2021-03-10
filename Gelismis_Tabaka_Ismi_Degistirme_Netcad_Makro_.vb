
Sub Main
 Dim i,BD,BD2
 dim sagul,sagul2
 dim sabangul1,sabangul2,sabangul3,sabangul4,sabangul5
 dim sabangul,elifyaren
 with Netcad

set BD2 = Netcad.NewBDialog("Tabaka İsmi Değiştirme [Harita Akademi, Şaban GÜL]")



BD2.PutPrompt "Yapılan İşlemde Öncelikle Tabaka İsminin başına"
 BD2.PutPrompt "Sonra Sonuna, en son Bul Değiştir işlemleri yapılır. "
 BD2.PutPrompt "Metin Ekleme en fazla 3 karakter sınırındadır"
 BD2.PutPrompt " "
 BD2.PutPrompt "NETCADDE TABAKA ADI EN FAZLA 20 KARAKTERDİR."
 BD2.PutPrompt " "

if BD2.showmodal then
 else
 exit sub
 end if



set BD = Netcad.NewBDialog("Gelişmiş Nokta İsmi Değiştirme [Harita Akademi, Şaban GÜL]")

BD.PutPrompt "Tabaka Adının Başına ve Sonuna Metin Ekleme"
 BD.Getstring "elifyaren1","Başına Ekle:","", 3
 BD.Getstring "elifyaren2","Sonuna Ekle:","", 3
 BD.PutPrompt " "
 BD.PutPrompt "Tabaka Adında Bul Değiştir Yapma"
 BD.Getstring "elifyaren3","Eski Metin:","", 50
 BD.Getstring "elifyaren4","Yeni Metin:","", 50
 BD.PutPrompt " "





if BD.showmodal then
 sabangul1= BD.ValueByName("elifyaren1")
 sabangul2= BD.ValueByName("elifyaren2")
 sabangul3= BD.ValueByName("elifyaren3")
 sabangul4= BD.ValueByName("elifyaren4")



for i= 0 to .NumLayers-1

with nclayermanager
 elifyaren=.Layer(i).name

sabangul= sabangul1 & elifyaren
 sabangul=sabangul & sabangul2
 sabangul=replace(sabangul,sabangul3,sabangul4)
 .Layer(i).name = sabangul

end with



next
 end if

end with

End Sub
 
