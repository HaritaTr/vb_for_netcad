
sub main
 dim secim,c,layerno,obj,yaz1,yaz2,yaz3,yaz4,top,o,sagul,sel,i,j,elifyaren ,yazi,ruhan,gulailesi ,saban,kop
 Dim BD
 with netcad
 on error resume next
 set BD = Netcad.NewBDialog(" Seçilen Objedeki Yazıları Toplama [Harita Akademi, Şaban GÜL]")

BD.Getfloat "item1","Toplama Başlanacak İlk Değer: ",0,5
 BD.Getradio "item2","Proje Üzerindeki Seçim Türü","Tek Tek Seçim|Toplu Seçim" ,1
 BD.Getradio "item3","Seçilenleri SAGUL_YAZITOPLA Tabakasına Al","Evet|Hayır" ,1
 BD.Getradio "item4","Sonuç Gösterimi","Sadece Mesaj Olarak Göster|Sonucu Ekrada Yazdır|Hem Mesaj Göster Hem Ekrana Yazdır" ,0

if BD.showmodal then
 sagul=1
 elifyaren=1



sagul= BD.ValueByName("item2")
 elifyaren= BD.ValueByName("item3")
 ruhan= BD.ValueByName("item4")
 gulailesi=0
 if sagul=0 then
 set secim = .NewSelectStatus()
 set c = .newc(0,0,0)

top= BD.ValueByName("item1")
 set yazi = .NewSelectStatus

while .SelectObjectInstant ("Toplam yapılacak yazıları seçiniz. İşlem tamamlandıktan sonra fare ile sağ tıklayınız",1,array(otext),yazi)
 set o = yazi.objects(0)
 kop=top
 top=top+o.s

if kop<>top then
 gulailesi=gulailesi+1
 end if

if elifyaren=1 then
 o.tabaka=.createlayer("sagul_yazitopla",2)
 end if

.PutObject yazi.indexs(0), o
 wend
 End if



if sagul=1 then

top= BD.ValueByName("item1")
 set SEL = .NewSelectionSet ' Yeni kume yarat
 set o = .NewObject
 if SEL.SELECT("Toplam Yapılacak Yazıları Seçiniz",array(otext)) then ' istenen turleri kumeye ekle
 for i = 0 to SEL.NE-1 ' kumenin her bir elemani icin
 j = SEL.GetSelectedObject(i, o) ' objeyi ve gercek indeksini al
 kop=top
 top=top+o.s



if kop<>top then

gulailesi=gulailesi+1
 end if



if elifyaren=1 then
 o.tabaka=.createlayer("sagul_yazitopla",2)
 end if



.putobject j, o ' objeyi geri koy
 next
 SEL.RedrawAndRewind ' secim kumesini toplu kendi renginde
 end if ' cizdir ve kumeyi basa sardir.
 set SEL = nothing
 set o = nothing



end if

set c = nothing
 set yazi = nothing
 set o = nothing

if ruhan=0 or ruhan=2 then
 saban= round(top/gulailesi ,2)
 msgbox ("Seçilen Değerler Toplamı: "& top &chr(13)&chr(10)&" Toplam " & gulailesi & " tane seçildi." &chr(13)&chr(10)&" Ortalama= " & saban ),64,"Harita Akademi, Şaban GÜL"



end if

if ruhan=1 or ruhan=2 then
 set secim = .NewSelectStatus()
 set c = .newc(0,0,0)
 set yazi = .NewSelectStatus

if .SelectPoint("Nokta seç", c, 2) then
 set yaz1=.maketext(c,top,0,0,2,0,"1",.createlayer("ROL_CEPHE_1",4))
 .addobject(yaz1)
 end if

eND if

set o = nothing



eND if
 end with
 end sub
