
Sub Main

 Dim sagultabaka ,fk
 with NCLayerManager

.Add "@SAGUL",6
 .Add "SAGUL_NOKTA", 32
 .Add "SAGUL_ALAN", 48
 .Add "SAGUL_TABLO", 24
 end with




with netcad
  .SetParam PNC_PLINENAME, 1
  .SetParam PNC_PLINEHATCHS, 0
  .SetParam PNC_PNTNAMES, 1

Const ForReading = 1, ForWriting = 2, ForAppending = 8
 Dim fso, objFolder,f

Set fso = CreateObject("Scripting.FileSystemObject")
 Set objFolder = FSO.GetFolder("C:\")
 If Not FSO.FolderExists("C:\Sagul") Then
 objFolder.SubFolders.Add "Sagul"
 End If

Set fso = CreateObject("Scripting.FileSystemObject")
 Set objFolder = FSO.GetFolder("C:\Sagul")
 If Not FSO.FolderExists("C:\Sagul\Netcad") Then
 objFolder.SubFolders.Add "Netcad"
 End If

Set fso = CreateObject("Scripting.FileSystemObject")
 Set objFolder = FSO.GetFolder("C:\Sagul\Netcad")
 If Not FSO.FolderExists("C:\Sagul\Netcad\Makro") Then
 objFolder.SubFolders.Add "Makro"
 End If



Set fso = CreateObject("Scripting.FileSystemObject")
 Set objFolder = FSO.GetFolder("C:\Sagul\Netcad\Makro")
 If Not FSO.FolderExists("C:\Sagul\Netcad\Makro\Tanimlar") Then
 objFolder.SubFolders.Add "Tanimlar"
 End If


dim dosyaoku,frs ,eyups
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set frs = fso.OpenTextFile("C:\Sagul\Netcad\Makro\Tanimlar\jsonoku.sagul", ForReading, True)
       on error resume next


Do While Not frs.AtEndOfStream
 dosyaoku= Split(frs.ReadLine,",")
 Loop




  dim xfsot,xft
  if InStr(dosyaoku(0), ",") <> 0 then
  MsgBox "İlk Kurulum Hatası Oluştu. Sizin İçin Otomatik Düzelteceğiz",0,"Dikkat"
   Set xfsot = CreateObject("Scripting.FileSystemObject")
  Set xft = xfsot.OpenTextFile("C:\Sagul\Netcad\Makro\Tanimlar\jsonoku.sagul", ForWriting, True)

    xft.WriteLine ( 2 & "," & 2 & "," & 1 )
    xft.close
    refresh
    end if


 frs.Close

dim sgul1,sgul2,sgul3,sgul4,sgul5
 sgul1= dosyaoku(0)
 sgul2= dosyaoku(1)
 sgul3= dosyaoku(2)

dim BD,XLSpath,BD_SAGUL,tabaka,hakademi1,hakademi2,hakademi3,hakademi4 ,ha1,ha2,ha3,ha4



hakademi1= "SAGUL_NOKTA"
 hakademi2= "SAGUL_ALAN"
 hakademi3= "SAGUL_TABLO"
 hakademi4= ""

for tabaka = 0 to .numlayers - 1
 if .LayerNameOf(tabaka)=hakademi1 then ha1=tabaka
 next



for tabaka = 0 to .numlayers - 1
 if .LayerNameOf(tabaka)=hakademi2 then ha2=tabaka
 next



for tabaka = 0 to .numlayers - 1
 if .LayerNameOf(tabaka)=hakademi3 then ha3=tabaka
 next



for tabaka = 0 to .numlayers - 1
 if .LayerNameOf(tabaka)=hakademi4 then ha4=tabaka
 next



set BD = Netcad.NewBDialog("PARSEL SORGU DOSYASI AKTARMA 2. Sürüm ,[Harita Akademi, Şaban GÜL]")
 set BD_SAGUL = Netcad.NewBDialog("PARSEL SORGU DOSYASI AKTARMA , [Harita Akademi, Şaban GÜL]")
 BD.PutPrompt "DİKKAT:Lütfen www.sabangul.com.tr adresinden uyarıları okuyunuz"
 BD.PutPrompt " "
 BD.GetFileName "sabangul","TKGM Parsel Sorgu Json Dosyasını Seçiniz...","","Json Dosyası|*.json|Tum Dosyalar|*.*","xls"

BD.GetCombo "item1","Dönüştürülecek Koordinat Sistemi (UTM3)","Dönüşüm Yapma|ED50|ITRF" ,sgul1
 BD.GetCombo "item2","Dönüştürülecek Dilim Numarası","27|30|33|36|39|42|45" ,sgul2
BD.GetCombo "item3","Koordinat Dönüşümünde Projede Dönüşsün mü ? ","Hayır|Evet" ,sgul3
 if BD.showmodal then
        xlspath = BD.ValueByName("sabangul")
            Dim fsot, ft
            Set fsot = CreateObject("Scripting.FileSystemObject")
            Set ft = fsot.OpenTextFile("C:\Sagul\Netcad\Makro\Tanimlar\jsonoku.sagul", ForWriting, True)
            dim gul1,gul2,gul3,gul4
            gul1= BD.ValueByName("item1")
            gul2= BD.ValueByName("item2")
            gul3= BD.ValueByName("item3")
            ft.WriteLine ( gul1 & "," & gul2 & "," & gul3 )
            ft.close
            dim ahmet,furkan,uncu
             ahmet= 0
             furkan= 33
             uncu=0
             if BD.ValueByName("item1")=0 then ahmet=0
             if BD.ValueByName("item1")=1 then ahmet=4
             if BD.ValueByName("item1")=2 then ahmet=1
             if BD.ValueByName("item2")=0 then furkan=27
             if BD.ValueByName("item2")=1 then furkan=30
             if BD.ValueByName("item2")=2 then furkan=33
             if BD.ValueByName("item2")=3 then furkan=36
             if BD.ValueByName("item2")=4 then furkan=39
             if BD.ValueByName("item2")=5 then furkan=42
             if BD.ValueByName("item2")=6 then furkan=45
             if BD.ValueByName("item3")=1 then uncu=1
 else
 exit sub
 end if












if XLSpath=-1 or XLSpath="" then
 msgbox "Dosya Seçilmedi"
exit sub
 end if





Dim Satir ,o ,a ,c1,c2,c3,c4,c5,c6

Set fso = CreateObject("Scripting.FileSystemObject")
 Set f = fso.OpenTextFile(BD.ValueByName("sabangul"), ForReading, True)



set BD = Nothing

  
  hakademi1= "SAGUL_ALAN"
 hakademi2= "SAGUL_NOKTA"
 hakademi3= "SAGUL_TABLO"
 hakademi4= ""

a= F.ReadLine
 a= replace(a,"[[[","$")
 a= replace(a,"]]]","$")


if ahmet<>0 then
 dim pc
 set pc = Netcad.NewProjection
 pc.ProjectionType =1
 pc.Zone = 36
 pc.Datum = 0
 pc.SetToCurrentProject FALSE

end if



dim elif,yaren,gul,ruhan
 satir=Split(a,"$")
 elif = satir(0)
 yaren= satir(1)
 gul= satir(2)
 ruhan=""



dim saban,soner,olcayto,sagul1,sagul2
 saban=""""
 soner="},properties:{"
 olcayto="}}],type:FeatureCollection,crs:{type:name,properties:{name:EPSG:4326}}}"
 sagul1=replace(yaren,"],[","$")

sagul2 =Replace(gul,saban,"")
 sagul2 =Replace(sagul2,soner,"")
 sagul2 =Replace(sagul2,olcayto,"")
 sagul2 =Replace(sagul2,","," ")



dim sagulnet,sagulnet1,sagulnet2,sagul_1,sagul_2,sagul_3,sagul_4,sagul_5,nitelik
 sagulnet1=Split(sagul2,"Alan")
 sagulnet2=Split(sagul2,"m2")
 sagul_1=replace(sagulnet2(0),sagulnet1(0),"")
 sagul_2=replace(sagul_1,"Alan:","")
 sagul_3=replace(sagul_2," ","")

sagul_4 = replace(sagul2,sagul_2,sagul_3)
 sagul_5 = replace(sagul_4,"m2","")

sagulnet=replace(sagul_5,"ParselNo","")
 sagulnet=replace(sagulnet,"Alan","")
 sagulnet=replace(sagulnet,"Mevkii","")
 sagulnet=replace(sagulnet,"Nitelik","")
 sagulnet=replace(sagulnet,"Ada","")
 sagulnet=replace(sagulnet,"Ilce","")
 sagulnet=replace(sagulnet,"Il","")
 sagulnet=replace(sagulnet,"Pafta","")
 sagulnet=replace(sagulnet,"Mahalle","")

nitelik=Split(sagulnet," :")

dim n1,n2,n3,n4,n5,n6,n7,n8,n9
 n1= replace(nitelik(0),":","") ' parsel no
 n2= nitelik(1) ' Alan
 n3= nitelik(2) ' Mevkii
 n4= nitelik(3) ' Nitelik
 n5= replace(nitelik(4)," ","") ' Ada
 n6= nitelik(5) ' Il
 n7= nitelik(6) ' Ilce
 n8= nitelik(7) ' Pafta
 n9= nitelik(8) ' Mahalle

dim ili,ilcesi,mahallesi,adaparseli,alani,mevkiisi,paftasi,cinsi

ili= n6
 ilcesi= n7
 mahallesi= n9
 adaparseli= n5 & "_" & n1
 alani= n2
 mevkiisi= n3
 paftasi= n8
 cinsi= n4
 paftasi= n8

dim x,kubilay ,satir2
 dim i,j,p
 dim yasin
 dim soner1,soner2
 Satir=Split(sagul1,"$")
 kubilay=0

set p = nothing
 set p = .NewPoly
 for each x in satir
 kubilay=kubilay+1
 Satir2=Split(x,",")
 soner1= satir2(0)
 soner2= satir2(1)
 p.AddCoor(.NewC(soner1,soner2,0))



.AddObject .MakePoint(.newc(soner1,soner2,0), kubilay,"SAGULNET" ,ha1)
 next
 set o = .MakePline(adaparseli,POLYCLOSED+POLYFILLED+16,alani,ha2,0,0,p)
 .AddObject o



set o = nothing
 set p = nothing



f.Close





if ahmet<>0 then

set pc = Netcad.NewProjection
 pc.ProjectionType =3
 pc.Zone = furkan
 pc.Datum = ahmet
 if uncu=1 then
 pc.SetToCurrentProject true
 else
 pc.SetToCurrentProject false
 end if
 end if







.findworld







dim secim,c,obj,yazi,layerno ,yaz1,yaz2,yaz3,yaz4,yaz5,yaz6,yaz7,yaz8,yaz9,yaz10,asilyazi,yaziboy,koorx
 set secim = .NewSelectStatus()
 set c = .newc(0,0,0)






asilyazi=2
 yaziboy=3
 layerno=.foundlayer("TXT_KOORDINAT")

if .SelectPoint("Tablonun Yerleştirileceği Yeri Seçiniz", c, 2) then

koorx=c.x
 koorx=koorx+2

c.x=koorx-(yaziboy)
 set yaz1=.maketext(c, adaparseli & " No'lu Parsel Bilgileri",0,0,asilyazi,0,"1",ha3)
 c.x=koorx-(yaziboy*2)
 set yaz2=.maketext(c, "İli: " & ili,0,0,asilyazi,0,"1",ha3)
 c.x=koorx-(yaziboy*3)
 set yaz3=.maketext(c, "İlçesi: " & ilcesi,0,0,asilyazi,0,"1",ha3)
 c.x=koorx-(yaziboy*4)
 set yaz4=.maketext(c, "Mahallesi: " & mahallesi,0,0,asilyazi,0,"1",ha3)
 c.x=koorx-(yaziboy*5)
 set yaz5=.maketext(c, "Alanı: " & alani,0,0,asilyazi,0,"1",ha3)
 c.x=koorx-(yaziboy*6)
 set yaz6=.maketext(c, "Cinsi: " & cinsi,0,0,asilyazi,0,"1",ha3)
 c.x=koorx-(yaziboy*7)
 set yaz7=.maketext(c, "Paftası: " & paftasi,0,0,asilyazi,0,"1",ha3)
 c.x=koorx-(yaziboy*8)
 set yaz8=.maketext(c, "Mevkiisi: " & mevkiisi,0,0,asilyazi,0,"1",ha3)



.addobject(yaz1)
 .addobject(yaz2)
 .addobject(yaz3)
 .addobject(yaz4)
 .addobject(yaz5)
 .addobject(yaz6)
 .addobject(yaz7)
 .addobject(yaz8)

'.AddObject .MakeLine(.newc(c.y,c.x+(asilyazi*3),0), .newc(c.y+asilyazi*25,c.x+(asilyazi*3),0), .CreateLayer("TABAKA1",yellow), 0, 0)


end if







End With




 





End Sub
