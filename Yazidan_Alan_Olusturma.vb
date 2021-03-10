
SUB Main
Dim SAGUL_DLG,i,alanlar,yazilar,parsel,YaziYazdir,ruhan,elif,sabangul
with netcad
set SAGUL_DLG = .NewBDialog("Alan İçindeki Yazıları Alan Adı Yapma [Şaban GÜL, Harita Mühendisi]")
SAGUL_DLG.GetCombo"sagul1","Alanlar Hangi Tabakada ?", "",1
for i = 0 to .numlayers - 1
SAGUL_DLG.AddCombo .LayerNameOf(i)
next
SAGUL_DLG.GetCombo"sagul2","Yazılar Hangi Tabakada ?", "", 1
for i = 0 to .numlayers - 1
SAGUL_DLG.AddCombo .LayerNameOf(i)
next
SAGUL_DLG.Getcheck "sagul3","Alan Adı Değişenleri 0_DEGISEN tabakasına al", 0
if SAGUL_DLG.showmodal then
alanlar=SAGUL_DLG.ValueByName("sagul1")
yazilar=SAGUL_DLG.ValueByName("sagul2")

if SAGUL_DLG.ValueByName("sagul3") = 1 then
with nclayermanager
.add "0_DEGISEN",4
end with
end if
 sabangul= SAGUL_DLG.ValueByName("sagul3")
if alanlar="" then
Msgbox "Alan Tabakası Bulunamadı",0,"Dikkat!"
Exit Sub
End if
if yazilar="" then
Msgbox "Yazı Tabakası Bulunamadı" ,0,"Dikkat!"
Exit Sub
End if
' www.sabangul.com.tr Web Sayfasından İndirilmiştir
' Şaban GÜL , Harita Mühendisi
' Her Türlü Hata, İstek ve Öneriler İçin
' haritaakademi@gmail.com veya sagulnet@gmail.com
' adresine durumu anlatan bir e-posta gönderiniz.

set SAGUL_DLG = Nothing
.SetFilter nothing, array(alanlar), array(opline)
for i = 0 to .numobject-1
.BackMessage
set parsel = .getobject(i) ' i. objeyi al
if parsel.tabaka=alanlar and parsel.tag = opline then 'parsel.tag = opline and 
YaziYazdir= parsel.pname
.SetMessage YaziYazdir
.SetFilter .ObjectExtends(Parsel), array(yazilar), array(otext) 
Do
set ruhan= .GetNextObject
if ruhan is nothing then
exit do
else
set elif=.GetPlineExt(parsel)
if elif.InPoly(ruhan.p1) then
parsel.pname=ruhan.s
if sabangul = 1 then
parsel.tabaka= .foundlayer("0_DEGISEN")
end if
.PutObject i, parsel
end if
end if
set ruhan = nothing 
Loop
.ResetFilter
end if
set parsel=nothing
next
.ResetFilter

msgbox"İşlem Başarıyla Tamamlanmıştır",0,"www.sabangul.com.tr"
end if
end with

END SUB
