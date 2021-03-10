
Sub Main
Dim i
dim obj
dim layerno
dim BD
dim sagul

set BD = Netcad.NewBDialog("GRID VE GRIDKOOR Tabakalarının ismine datum ismi ekle,[Harita Akademi, Şaban GÜL]")

BD.Getradio "item1","Eklenecek Metin Giriniz","ED50|ITRF|MEVZII" ,0
 BD.GetString "item","Bir Ayraç Giriniz","_",1
 BD.PutPrompt " "
 BD.PutPrompt " DİKKAT!"&chr(13)&chr(10)&" GRID ve GRIDKOOR tabakalarının ismi değişecek"
 BD.PutPrompt " "



if BD.showmodal then
 sagul=""
 if BD.ValueByName("item1")=0 then sagul="ED50"
 if BD.ValueByName("item1")=1 then sagul="ITRF"
 if BD.ValueByName("item1")=2 then sagul="MEVZII"

on error resume next

with NCLayerManager
 .Layer( .Find("GRID") ).name= "GRID" & BD.ValueByName("item") & sagul
 .Layer( .Find("GRIDKOOR") ).name= "GRIDKOOR" & BD.ValueByName("item") & sagul
 .layer(.Find("GRIDKOOR_"&sagul)).color=17
 .layer(.Find("GRID_"&sagul)).color=17
 end with

END IF



with Netcad

.SetFilter nothing, array(.foundlayer("GRIDKOOR_ED50")), array(otext)

do
 set obj=.getnextobject
 if obj is nothing then
 exit do
 end if

obj.flags=1
 obj.s="("&obj.s&")"



.PUTOBJECT .CUROBJPOS,OBJ



loop





end with
dim elif1,elif2
elif1= "GRID" & BD.ValueByName("item") & sagul
elif2= "GRIDKOOR" & BD.ValueByName("item") & sagul
msgbox ("İşlem Başarıyla Tamamlandı." &chr(13)&chr(10)&"GRID Tabakasının İsmi " & elif1 & chr(13)&chr(10)&" GRIDKOOR Tabakasının İsmi " & elif2 & " olmuştur." ),64,"Harita Akademi, Şaban GÜL"

End Sub
