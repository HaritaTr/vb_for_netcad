
Sub Main
 with netcad



Dim i
 dim j
 dim o
 dim SEL
 dim xls
 dim xlspath
 dim alan
 dim DEG
 dim CL
 dim bd
 DIM U,V,R,W
 dim ruhangul
 dim elifyaren
 DIM NO(50000,2)
 DEG = ""
 CL=0:U=0:R=0
 set xls = CreateObject("excel.application")







set BD = Netcad.NewBDialog("Excelden Kritere Göre Tabakalandýrma [Harita Akademi, Þaban GÜL]")
 BD.GetFileName "item1","Aktarým Yapýlacak Excel Dosyasý Seçiniz:","","Excel Dosyalari|*.xls|Tum Dosyalar|*.*","xls"
 BD.Getcombo "item2","Parsel Numarasý Hangi Sütunda Bulunuyor ? ","A|B|C|D|E|F|G|H|I|J|K|L|M|N|O|P|Q|R|S|T|U|V|W|X|Y|Z" ,0
 BD.Getcombo "item3","Tabakalandýrýlacak Kriter Hangi Sütunda Bulunuyor ? ","A|B|C|D|E|F|G|H|I|J|K|L|M|N|O|P|Q|R|S|T|U|V|W|X|Y|Z" ,1
 BD.Getcheck "item4","Tabakalarý Gruplandýr.(@...)" ,1
 BD.GetString "item5", "Grup Adý", "sagul", 5
 BD.GetString "item6", "Boþ Deðerlerin Alýnacaðý Tabaka", "BOÞ_DEÐER", 10
 if BD.showmodal then
 xlspath = BD.ValueByName("item1")
 RUHANGUL=BD.ValueByName("item4")
 elifyaren= BD.ValueByName("item5")
 else
 exit sub
 end if

dim saban,ruhan

saban= BD.ValueByName("item2")
 ruhan= BD.ValueByName("item3")

saban=1
 ruhan=2

if BD.ValueByName("item2")="A" then saban=1
 if BD.ValueByName("item2")="B" then saban=2
 if BD.ValueByName("item2")="C" then saban=3
 if BD.ValueByName("item2")="D" then saban=4
 if BD.ValueByName("item2")="E" then saban=5
 if BD.ValueByName("item2")="F" then saban=6
 if BD.ValueByName("item2")="G" then saban=7
 if BD.ValueByName("item2")="H" then saban=8
 if BD.ValueByName("item2")="I" then saban=9
 if BD.ValueByName("item2")="J" then saban=10
 if BD.ValueByName("item2")="K" then saban=11
 if BD.ValueByName("item2")="L" then saban=12
 if BD.ValueByName("item2")="M" then saban=13
 if BD.ValueByName("item2")="N" then saban=14
 if BD.ValueByName("item2")="O" then saban=15
 if BD.ValueByName("item2")="P" then saban=16
 if BD.ValueByName("item2")="Q" then saban=17
 if BD.ValueByName("item2")="R" then saban=18
 if BD.ValueByName("item2")="S" then saban=19
 if BD.ValueByName("item2")="T" then saban=20
 if BD.ValueByName("item2")="U" then saban=21
 if BD.ValueByName("item2")="V" then saban=22
 if BD.ValueByName("item2")="W" then saban=23
 if BD.ValueByName("item2")="X" then saban=24
 if BD.ValueByName("item2")="Y" then saban=25
 if BD.ValueByName("item2")="Z" then saban=26



if BD.ValueByName("item3")="A" then ruhan=1
 if BD.ValueByName("item3")="B" then ruhan=2
 if BD.ValueByName("item3")="C" then ruhan=3
 if BD.ValueByName("item3")="D" then ruhan=4
 if BD.ValueByName("item3")="E" then ruhan=5
 if BD.ValueByName("item3")="F" then ruhan=6
 if BD.ValueByName("item3")="G" then ruhan=7
 if BD.ValueByName("item3")="H" then ruhan=8
 if BD.ValueByName("item3")="I" then ruhan=9
 if BD.ValueByName("item3")="J" then ruhan=10
 if BD.ValueByName("item3")="K" then ruhan=11
 if BD.ValueByName("item3")="L" then ruhan=12
 if BD.ValueByName("item3")="M" then ruhan=13
 if BD.ValueByName("item3")="N" then ruhan=14
 if BD.ValueByName("item3")="O" then ruhan=15
 if BD.ValueByName("item3")="P" then ruhan=16
 if BD.ValueByName("item3")="Q" then ruhan=17
 if BD.ValueByName("item3")="R" then ruhan=18
 if BD.ValueByName("item3")="S" then ruhan=19
 if BD.ValueByName("item3")="T" then ruhan=20
 if BD.ValueByName("item3")="U" then ruhan=21
 if BD.ValueByName("item3")="V" then ruhan=22
 if BD.ValueByName("item3")="W" then ruhan=23
 if BD.ValueByName("item3")="X" then ruhan=24
 if BD.ValueByName("item3")="Y" then ruhan=25
 if BD.ValueByName("item3")="Z" then ruhan=26





set BD = Nothing
 xls.workbooks.open(xlspath)
 xls.range("A1").select

FOR U=1 TO 100000
 CL=CL+1
 NO(U,1)="*" & XLS.CELLS(U,saban)
 NO(U,2)= XLS.CELLS(U,ruhan)
 IF NO(U,2)="" THEN NO(U,2)=0
 IF NO(U,1)="*" THEN U=100000
 NEXT
 xls.quit





MSGBOX CL-1 & " Adet Parsel Excel Dosyasýndan Baþarýyla Okundu. Lütfen Ýþlem Görecek Parselleri Seçiniz."



FOR V=1 TO CL
 Dim max,min,elif ,sabangul,tabaka
 max=255
 min=1
 Randomize
 elif=Int((max-min+1)*Rnd+min)



if elif=0 or elif=15 then elif=4
 if elif>79 and elif <96 then elif=5

with NCLayerManager



if RUHANGUL=0 then
 .Add NO(V,2), elif
 else
 .Add "@"&elifyaren,5
 .Add elifyaren & "_" & NO(V,2), elif
 end if
 END With

sabangul=0
 for tabaka = 0 to .numlayers - 1
 sabangul=sabangul+1
 next
 if sabangul>254 then
 msgbox ("Tabaka Sayýsý Netcad'in Sýnýrýný Aþmak Üzere!!" &chr(13)&chr(10)&" Lütfen projenizi inceleyiniz veya tabakalarý azaltýnýz" ),64,"Harita Akademi, Þaban GÜL"
 msgbox ("Proje ve Veri Güvenliði Ýçin Ýþleme Devam Edilmeyecektir." &chr(13)&chr(10)&"Tabakalarý azaltýp tekrar deneyiniz." ),64,"Harita Akademi, Þaban GÜL"
 exit sub
 end if

next

with Netcad
 set SEL = .NewSelectionSet
 set o = .NewObject





if SEL.SELECT("Lütfen Ýþlem Görecek Parselleri Seçiniz...",array(opline)) then

for i = 0 to SEL.NE-1
 j = SEL.GetSelectedObject(i, o)
 alan = o.pname

FOR V=1 TO CL
 W=NO(V,1)
 if W ="*" & alan then
 with NCLayerManager
 if RUHANGUL=0 then
 o.tabaka = .Find(NO(V,2))
 else
 o.tabaka = .Find(elifyaren & "_" & NO(V,2))
 end if

if .Find(NO(V,2))="" then
 o.tabaka=BD.ValueByName("item6")
 end if

end with
 .putobject j, o
 R=R+1
 V=U
 end if
 NEXT
 next
 SEL.RedrawAndRewind

end if



set SEL = nothing
 set o = nothing
 end with
 end with
 MSGBOX R & " adet Parselin Tabakasý Deðiþtirildi."
 end sub
 
