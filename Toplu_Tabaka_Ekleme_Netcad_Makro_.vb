
Sub Main
 Dim i
 Dim Bd ,bd0,bd1,bd2
 dim sagul ,ruhan
 Dim Satir ,c1 ,o ,SGL
 Const ForReading = 1, ForWriting = 2, ForAppending = 8
 dim fso,f,listx ,fi
 Dim Line
 With Netcad

with NCLayerManager



set BD = Netcad.NewBDialog("Toplu Tabaka Adı Girişi")
 BD.Getradio "item1","Bir Yöntem Seçiniz","@TANIM Girip Bağlı 10'a Kadar Tabaka Girişi|Tanımlı Tabaka Grubu Girişi|Not Defterinden Alım|LYR Dosyasından Alım" ,0



if BD.SHOWMODAL then
 sagul=BD.ValueByName("item1")
 else
 exit sub
 end if



if sagul=0 then
 set BD0 = Netcad.NewBDialog("Toplu Tabaka Adı Girişi [Harita Akademi, Şaban GÜL]")
 BD0.GetString "item","Grup Adı Giriniz: (Enfazla 9 karakter)","SAGUL",9
 BD0.PutPrompt "Yukarıda Girdiniğiniz Grup @ ile grup yapılacak ve aşağıdakiler yazılacak."
 BD0.PutPrompt "Örneğin KAD yazdığınızda @KAD açılıp @KAD_TAB1, @KAD_TAB2 .. açılır"
 BD0.PutPrompt "______________________________________________________________"
 BD0.GetString "item1","1. Tabaka Adı","",10
 BD0.GetString "item2","2. Tabaka Adı","",10
 BD0.GetString "item3","3. Tabaka Adı","",10
 BD0.GetString "item4","4. Tabaka Adı","",10
 BD0.GetString "item5","5. Tabaka Adı","",10
 BD0.GetString "item6","6. Tabaka Adı","",10
 BD0.GetString "item7","7. Tabaka Adı","",10
 BD0.GetString "item8","8. Tabaka Adı","",10
 BD0.GetString "item9","9. Tabaka Adı","",10
 BD0.GetString "item10","10. Tabaka Adı","",10

if BD0.SHOWMODAL then
 .Add "@" & BD0.ValueByName("item"),3
 if BD0.ValueByName("item1")<>"" then .Add BD0.ValueByName("item1"),1
 if BD0.ValueByName("item2")<>"" then .Add BD0.ValueByName("item2"),2
 if BD0.ValueByName("item3")<>"" then .Add BD0.ValueByName("item3"),3
 if BD0.ValueByName("item4")<>"" then .Add BD0.ValueByName("item4"),4
 if BD0.ValueByName("item5")<>"" then .Add BD0.ValueByName("item5"),5
 if BD0.ValueByName("item6")<>"" then .Add BD0.ValueByName("item6"),6
 if BD0.ValueByName("item7")<>"" then .Add BD0.ValueByName("item7"),7
 if BD0.ValueByName("item8")<>"" then .Add BD0.ValueByName("item8"),8
 if BD0.ValueByName("item9")<>"" then .Add BD0.ValueByName("item9"),9
 if BD0.ValueByName("item10")<>"" then .Add BD0.ValueByName("item10"),10
 else
 exit sub
 end if
 end if

if sagul=1 then
 set BD1 = Netcad.NewBDialog("Toplu Tabaka Adı Girişi [Harita Akademi, Şaban GÜL] ")
 BD1.PutPrompt "Yukarıda Girdiniğiniz Grup @ ile grup yapılacak ve aşağıdakiler yazılacak."
 BD1.Getradio "item1","Bir Yöntem Seçiniz","@KAD,KAD_ADA,KAD_PARSEL,KAD_PAFTA|@PLAN,@PLAN_KUZEY...|@KAMU,KAMU_A,KAMU_B,KAMU_C|@TRM,TRM_SAHIS,TRM_MALIYE,TRM_ORMAN|@2B,2B_NOKTA,2B_CIZGI" ,0
 if BD1.SHOWMODAL then
 if BD1.ValueByName("item1")=0 THEN
 .Add "@KAD",1
 .Add "KAD_ALAN",2
 .Add "KAD_CIZGI",3
 .Add "KAD_YAZI",4
 .Add "KAD_ADA",5
 .Add "KAD_PARSEL",6
 .Add "KAD_NO_ADA",7
 .Add "KAD_NO_PARSEL",8
 .Add "KAD_DETAY",9
 .Add "KAD_NADI",10
 .Add "KAD_MEVKII",11
 .Add "KAD_MALIK",12
 .Add "KAD_GS",12
 .Add "KAD_GS_NOKTA",12
 END if



if BD1.ValueByName("item1")=1 THEN
 .Add "@PLAN",1
 .Add "PLAN_KUZEY",2
 .Add "PLAN_CIZGI", 3
 .Add "PLAN_LEJANT", 4
 .Add "PLAN_ANTET", 5
 .Add "PLAN_ILILCE", 6
 .Add "PLAN_ACIKLAMA", 7
 .Add "PLAN_TABLO", 8
 .Add "PLAN_YAZI", 9
 END if

if BD1.ValueByName("item1")=2 THEN
 .Add "@KAMU",1
 .Add "KAMU_ALAN",2
 .Add "KAMU_KOMSU", 3
 .Add "KAMU_SINIR", 4
 .Add "KAMU_A", 5
 .Add "KAMU_B", 6
 .Add "KAMU_C", 7
 .Add "KAMU_ESKI", 8
 .Add "KAMU_IRT", 9

END if

if BD1.ValueByName("item1")=3 THEN
 .Add "@TRM",1
 .Add "TRM_SAHIS",2
 .Add "TRM_MALIYE", 3
 .Add "TRM_ORMAN", 4
 .Add "TRM_MERA", 5
 .Add "TRM_YOL", 6
 .Add "TRM_DERE", 7
 .Add "TRM_ESKI", 8
 .Add "TRM_TUZEL", 9
 .Add "TRM_DAVALI", 10
 END if



if BD1.ValueByName("item1")=4 THEN
 .Add "@2B",1
 .Add "2B_CIZGI",2
 .Add "2B_NOKTA", 3
 .Add "2B_ALAN", 4
 .Add "2B_ELGPS", 5
 .Add "2B_ARAZI", 6
 .Add "2B_MALIK", 7
 .Add "2B_NOT", 8
 END if
 END if
 end if

if sagul=11 then
 Msgbox "Bu menü sadece Uzman Kullanıcı Grupları tarafından açılabilir."



end if

if sagul=2 then
 ' Msgbox "Bu menü sadece Uzman Kullanıcı Grupları tarafından açılabilir."
 set SGL = Netcad.NewBDialog("Dosya Seç [Harita Akademi, Şaban GÜL] ")
 SGL.GetFileName "item7","Text Dosyasını Seçiniz...","","Text Dosyalari|*.txt|Tum Dosyalar|*.*","txt"
 if SGL.showmodal then
 SAGUL = SGL.ValueByName("item7")
 else
 exit sub
 end if



Set fso = CreateObject("Scripting.FileSystemObject")
 Set f = fso.OpenTextFile(SAGUL, ForReading, True)
 Set listx = CreateObject("System.Collections.ArrayList")

Do While Not f.AtEndOfStream
 Line = f.readline
 listx.add line
 Loop
 f.Close
 for fi=0 to listx.count- 1
 with NCLayerManager
 .Add listx(fi),fi+1
 end with
 next
 end if

if sagul=3 then
 ' Msgbox "Bu menü sadece Uzman Kullanıcı Grupları tarafından açılabilir."
 set SGL = Netcad.NewBDialog("Dosya Seç... [Harita Akademi, Şaban GÜL] ")
 SGL.GetFileName "item7","LYR Dosyasını Seçiniz...","","LYR Dosyalari|*.LYR|Tum Dosyalar|*.*","LYR"
 if SGL.showmodal then
 SAGUL = SGL.ValueByName("item7")
 else
 exit sub
 end if



Set fso = CreateObject("Scripting.FileSystemObject")
 Set f = fso.OpenTextFile(SAGUL, ForReading, True)
 Set listx = CreateObject("System.Collections.ArrayList")

Do While Not f.AtEndOfStream
 Line = f.readline
 line=replace(line," ","")
 line=replace(line," ","")
 line=replace(line," ","")
 line=replace(line," ","")
 line=replace(line," ","")

if mid(line,1,11)="<LayerName>" then

line=replace(line,"<LayerName>","")
 line=replace(line,"</LayerName>","")
 listx.add line
 end if
 Loop
 f.Close
 for fi=0 to listx.count- 1
 with NCLayerManager
 .Add listx(fi),fi+1
 end with
 next



end if



end with
 end with



End Sub
