
Sub Main

with netcad

Dim obj
 dim selection
 Dim BD
 dim item



set BD = Netcad.NewBDialog("Yazý Boyunun Deðiþtirilmesi [Harita Akademi, Þaban GÜL]")

BD.Getfloat "item","Yeni Yazý Boyu:",2,1

if BD.showmodal then

set selection = .NewSelectStatus

while .SelectObjectInstant("BOYU DEÐÝÞECEK YAZI GRUBUNA AÝT BÝR YAZI OBJESÝ SEÇÝNÝZ",1,array(otext),selection)
 set obj = selection.objects(0)
 .DrawObject obj,blue



.SetFilter nothing, array(obj.tabaka), array(oText)

do

set obj=.getnextobject

if obj is nothing then

exit do

end if

obj.sc = BD.ValueByName("item")
 .PutObject .CurObjPos, obj

loop

.resetfilter

wend

end if

.netcadcommand("REGEN")



end with

end sub
