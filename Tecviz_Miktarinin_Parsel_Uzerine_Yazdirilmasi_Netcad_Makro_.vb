
Sub Main

with netcad
 Dim a,b,c,m,alan,o,d,e,poly,s,y
 Dim i, SGL_DLG,tabaka,tabaka2,tabaka3,yaziboy,tecvizx
 Dim Olcek,tapu,hesap,mfark,tecviz,tecviz1,tecviz2,tecdur
 olcek=.getparam(94)

set SGL_DLG = Netcad.NewBDialog("Tecviz Miktarlarının Yazdırılması [Harita Akademi, Şaban GÜL]")
 'SGL_DLG.PutPrompt "----------AŞAĞIDAKİ ALANDAN TECVİZ PARAMETRELERİNİ BELİRLEYİNİZ----------"

SGL_DLG.GetCombo "tabaka", "Hangi Tabakadakiler İçin Uygulansın", "Tüm Tabakalar", 0
 for i =0 to .numlayers - 1
 SGL_DLG.AddCombo .LayerNameOf(i)
 next
 SGL_DLG.Getfloat "sagulnet1","Ölçek Değeri (0 için proje ölçeği esas alınır )",olcek,0
 SGL_DLG.Getradio "sagulnet5","Tecviz Tipi Belirleyiniz","Yapılaşma Yok|Yapılaşma Var" ,0
 SGL_DLG.GETCHECK "sagulnet2","Tapu Alanı Yoksa Hesap Alanı Esas Alınsın",1
 SGL_DLG.PutPrompt " "
 'SGL_DLG.PutPrompt "----------SONUÇLAR İLE İLGİLİ PARAMETRELERİ BELİRLEYİNİZ----------"
 SGL_DLG.GetCombo "tabaka2", "Tecviz Hesabı Sonucu (Tecviz İçinde İse)", "0", 0
 for i =0 to .numlayers - 1
 SGL_DLG.AddCombo .LayerNameOf(i)
 next
 SGL_DLG.GetCombo "tabaka3", "Tecviz Hesabı Sonucu (Tecviz Dışında İse)", "0", 0
 for i =0 to .numlayers - 1
 SGL_DLG.AddCombo .LayerNameOf(i)
 next

SGL_DLG.Getfloat "sagulnet3","Tecvizin Değerinin Yazı Boyu",5,0
 SGL_DLG.GETCHECK "sagulnet4","Tecviz Yazı Boyu proje ölçeği ile çarpılsın",1
 SGL_DLG.GETCHECK "sagulnet6", "Tecviz Miktarının Sağına m² Yazısı Ekle", 1
 SGL_DLG.GETCHECK "sagulnet7", "Tecviz Dışında İse (D) Yazısı Ekle", 1
 SGL_DLG.GETCHECK "sagulnet8", "Tecviz Dışında İse Rengini Kırmızı Yap", 1
 if SGL_DLG.showmodal then
 yaziboy= SGL_DLG.ValueByName("sagulnet3")
 if yaziboy<1 then yaziboy=1
 if SGL_DLG.ValueByName("sagulnet4")=1 then yaziboy= round(yaziboy*olcek/1000,2)



with nclayermanager

if SGL_DLG.ValueByName("tabaka")<>0 then
 tabaka= .layer(SGL_DLG.ValueByName("tabaka")-1).name
 end if
 tabaka2 = .layer(SGL_DLG.ValueByName("tabaka2")).name
 tabaka3 = .layer(SGL_DLG.ValueByName("tabaka3")).name
 end with
 else
 exit sub
 end if





if tabaka="" then
 .setfilter nothing, array(),array(opline)
 else
 .setfilter nothing, array(tabaka),array(opline)
 end if





do
 set o=.getnextobject
 if o is nothing then exit do
 set poly=.getplineext(o)

if SGL_DLG.ValueByName("sagulnet1")<>0 THEN olcek=SGL_DLG.ValueByName("sagulnet1")
 tapu=o.tarea
 hesap=o.area
 if tapu=0 and SGL_DLG.ValueByName("sagulnet2")=1 then tapu=hesap
 mfark=abs(tapu-hesap)

tecviz1=0.0004*olcek*sqr(tapu)+0.0003*tapu
 tecviz2=0.013*sqr(tapu*olcek)+0.0003*tapu

if SGL_DLG.ValueByName("sagulnet5")=0 then
 tecviz=Tecviz1
 else
 tecviz=Tecviz2
 end if

tecviz=round(Tecviz,3)
 tecvizx=tecviz
 if SGL_DLG.ValueByName("sagulnet6")=1 then tecviz = tecviz &"m²"
 if SGL_DLG.ValueByName("sagulnet7")=1 and mfark> tecvizx then tecviz = tecviz &"(D)"

if mfark>tecvizx then
 set y=.MakeText(poly.centerofmass,tecviz, 0,1,yaziboy,0,0,.createlayer(tabaka2,1))
 if SGL_DLG.ValueByName("sagulnet8")=1 then y.renk=Red
 .addobject y

else
 set y=.MakeText(poly.centerofmass,tecviz, 0,1,yaziboy,0,0,.createlayer(tabaka3,1))

.addobject y

end if





loop
 end with
 end sub
