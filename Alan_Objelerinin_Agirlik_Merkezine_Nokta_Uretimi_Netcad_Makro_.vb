
Sub Main
 With netcad

Dim i,j,o,SEL,tx, p ,TabNo,BD1,BD2,BD3,BD4,tabakam

tabakam=""
 set BD1 = Netcad.NewBDialog("Hangi Tabakada Yazdırılacağını Seçiniz [Harita Akademi, Şaban GÜL]")
 set BD2 = Netcad.NewBDialog("Tabaka Seçimi Yapınız [Harita Akademi, Şaban GÜL]")
 set BD3 = Netcad.NewBDialog("Yeni Tabaka Adı Giriniz [Harita Akademi, Şaban GÜL]")
 set BD4 = Netcad.NewBDialog("Nokta Başına Veya Sonuna Eklenecek Metin [Harita Akademi, Şaban GÜL]")





BD1.GetCombo "tabaka", "Parsel No Hangi Tabakaya Yazılsın: ", "0", 0
 for i = 1 to .numlayers - 1
 BD1.AddCombo .LayerNameOf(i)
 next



BD2.Getradio "secim","Noktalar Hangi Tabakada Oluşturulsun?","Mevcut Tabakalar|Yeni Tabaka" ,0

BD3.Getstring "yenitabaka","Yeni Tabaka Adı: ","sagul_agirlik_merkezi",50



BD4.Getradio "metin","Noktanın Başı veya Sonuna Veri Ekleme","Herhangi İşlem Yapma ( Olduğu Gibi Bırak )|Noktanın Başına Aşağıdaki Metni Ekle|Noktanın Sonuna Aşağıdaki Metni Ekle| Noktadan Aşağıdaki Karakteri Sil" ,0
 BD4.Getstring "metin2","...","",50
 if BD4.showmodal then
 dim saban
 saban=BD4.ValueByName("metin2")
 if BD2.showmodal then

if BD2.ValueByName("secim")=0 then
 if BD1.showmodal then

tabakam=BD1.ValueByName("tabaka")
 with ncLayermanager
 tabakam= .layer(tabakam).Name

end with
 end if

else

if BD3.showmodal then
 tabakam= BD3.ValueByName("yenitabaka")
 end if

end if

if tabakam="" then
 exit sub
 end if



dim elif
 elif=0
 set SEL = .NewSelectionSet ' Yeni kume yarat
 set o = .NewObject
 set p=.Newpoly
 TabNo = Netcad.CreateLayer (tabakam, 4)
 if SEL.SELECT("Ağırlık Merkezi Noktası Oluşturulacak Alanları Seçiniz",array(opline)) then ' istenen turleri kumeye ekle
 for i = 0 to SEL.NE-1 ' kumenin her bir elemani icin
 elif=elif+1
 j = SEL.GetSelectedObject(i, o) ' objeyi ve gercek indeksini al
 set p=.getplineExt(o)
 o.renk = red ' rengini sari yap

dim ruhan
 ruhan=""
 ruhan = o.pname
 if BD4.ValueByName("metin")=0 then ruhan= ruhan
 if BD4.ValueByName("metin")=1 then ruhan= saban & ruhan
 if BD4.ValueByName("metin")=2 then ruhan= ruhan & saban
 if BD4.ValueByName("metin")=3 then ruhan= replace(ruhan,saban,"")
 set tx=.MakePoint(p.CenterOfMass,ruhan,0,TabNo)
 .addobject(tx)

next
 SEL.RedrawAndRewind ' secim kumesini toplu kendi renginde
 end if ' cizdir ve kumeyi basa sardir.
 set SEL = nothing
 set o = nothing

end if
 end if

END with

if elif=0 then
 msgbox ("Nokta üretimi yapılamadı"),64,"Harita Akademi, Şaban GÜL"
 else

msgbox ("Üretilen Nokta Sayısı: "& elif &chr(13)&chr(10)&" Noktalar " & tabakam & " tabakasına alındı." ),64,"Harita Akademi, Şaban GÜL"
 end if

end sub
