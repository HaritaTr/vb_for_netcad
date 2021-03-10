
Sub Main
 with netcad
 Dim a,b,c,m,alan,o,d,e,poly,s,y
 Dim i, SGL_DLG,tabaka,tabaka2,tabaka3,yaziboy,tecvizx
 Dim Olcek,tapu,hesap,mfark,tecviz,tecviz1,tecviz2,tecdur
 olcek=.getparam(94)

set SGL_DLG = Netcad.NewBDialog("Tecviz Miktarlarının Yazdırılması [Harita Akademi, Şaban GÜL]")

SGL_DLG.Getfloat "sagulnet1","Ölçek Değeri (0 için proje ölçeği esas alınır )",olcek,0
 SGL_DLG.Getradio "sagulnet5","Tecviz Tipi Belirleyiniz","Yapılaşma Yok|Yapılaşma Var" ,0
 'SGL_DLG.GETCHECK "sagulnet2","Tapu Alanı Yoksa Hesap Alanı Esas Alınsın",1



if SGL_DLG.showmodal then

end if





set alan = .NewSelectStatus
 while .SelectObjectInstant(" Tecviz Hesabı Yapılacak Alanı Seçiniz.[Harita Akademi, Şaban GÜL]",1,array(opline),alan)
 set o = alan.objects(0)

if SGL_DLG.ValueByName("sagulnet1")<>0 THEN olcek=SGL_DLG.ValueByName("sagulnet1")
 tapu=o.tarea
 hesap=o.area
 if tapu=0 and SGL_DLG.ValueByName("sagulnet2")=1 then tapu=hesap
 mfark=abs(tapu-hesap)

if tapu=0 then
 msgbox "Parsel Tapu Alanı Yoktur",64
 end if



tecviz1=0.0004*olcek*sqr(tapu)+0.0003*tapu
 tecviz2=0.013*sqr(tapu*olcek)+0.0003*tapu

if SGL_DLG.ValueByName("sagulnet5")=0 then
 tecviz=Tecviz1
 else
 tecviz=Tecviz2
 end if

tecviz=round(Tecviz,3)

e= abs(mfark-tecviz)

if mfark> tecviz and tapu > hesap then

msgbox "Parsel Tecviz Dışındadır"&chr(13)_
 & "Tapu Alanı=" & round(tapu,2) &chr(13)_
 & "Hesap Alanı=" & round(hesap,2) &chr(13)_
 & "Fark (Mutlak)=" & round(mfark,2) &chr(13)_
 & "Tecviz Miktarı=" & tecviz &chr(13)_
 &round(abs(e),2)&" m² ARTACAK",16

elseif mfark> tecviz and hesap > tapu then

msgbox "Parsel Tecviz Dışındadır"&chr(13)_
 & "Tapu Alanı=" & round(tapu,2) &chr(13)_
 & "Hesap Alanı=" & round(hesap,2) &chr(13)_
 & "Fark (Mutlak)=" & round(mfark,2) &chr(13)_
 & "Tecviz Miktarı=" & tecviz &chr(13)_
 &round(abs(e),2)&" m² AZALACAK",16

else

msgbox "Parsel Tecviz İçindedir"&chr(13)_
 & "Tapu Alanı=" & round(tapu,2) &chr(13)_
 & "Hesap Alanı=" & round(hesap,2) &chr(13)_
 & "Fark (Mutlak)=" & round(mfark,2) &chr(13)_
 & "Tecviz Miktarı=" & tecviz &chr(13),64



end if

wend

end with
 end sub
