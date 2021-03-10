
Sub Main
with netcad
Dim obj
dim selection
Dim BD
dim item



set BD = Netcad.NewBDialog("Yazıların Açısını Değiştirme,[Harita Akademi, Şaban GÜL]")

BD.Getfloat "item","Dönüklük Açısını Giriniz",0,1
 ' BD.Getradio "item1","Tanımlı Açı seçebilirsiniz","Yukarıdaki Açı|0|50|100|150|200|250|300|350|400" ,0



if BD.showmodal then

set selection = .NewSelectStatus

while .SelectObjectInstant("Dönecek Yazının Tabakasını Tanımlamak için bir tanesini seçiniz ",1,array(),selection)
 set obj = selection.objects(0)
 .DrawObject obj,blue



.SetFilter nothing, array(obj.tabaka), array()

do

set obj=.getnextobject

if obj is nothing then

exit do

end if

obj.angle = (BD.ValueByName("item"))*0.015707962
 .PutObject .CurObjPos, obj

loop

.resetfilter

wend

end if

.netcadcommand("REGEN")




end with

end sub
