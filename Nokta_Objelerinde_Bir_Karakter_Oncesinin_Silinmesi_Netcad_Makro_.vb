
sub main
with netcad
dim pos
dim obj
dim a
dim elif

dim BD
dim sabangul
dim elifyaren
 set BD = Netcad.NewBDialog("Nokta Objelerinin Adının Bir Kısmını Karaktere Göre Silme [Harita Akademi, Şaban GÜL]")

BD.Getstring "sagul1","Nokta Objesinin Adını Ayrıştıracak Karakteri Giriniz:","/" ,1
 ' BD.Getradio "sagul2","Hangi Kısmı Silinecek","Öncesi|Sonrası" ,0




 if BD.showmodal then
 else
 exit sub
 end if



if BD.ValueByName("sagul1")="" then
 msgbox ("Herhangi bir karakter girmediniz " &chr(13)&chr(10)&"Lütfen Bir Ayraç Giriniz" &chr(13)&chr(10)&"Örnek Karakter: /" ),64,"Harita Akademi, Şaban GÜL"
 end if

if BD.ValueByName("sagul2")=1 then
 elifyaren=1
 else
 elifyaren=0
 end if





.setfilter nothing,array(),array(opoint)
Do
set obj=.getnextobject
if obj is nothing then
exit do
else

pos=split(obj.s,"/")
a= obj.s
obj.pname=replace(obj.s,pos(0),"")

elif =Mid(obj.s,2,100000)
obj.s=elif
if obj.s="" then
else
 .putobject .curobjpos,obj
end if
end if

loop
end with
end sub
