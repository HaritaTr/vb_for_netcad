
Sub Main
 Dim obj,secim,layer,yazi,c,yaz,alan,obj1

with netcad
 set secim = .NewSelectStatus()
 .SelectObjectInstant "Kopyalanacak Yazı Seçiniz, [Harita Akademi]",1,array(otext),secim

set obj = secim.objects(0)
 .DrawObject obj,red
 yazi = obj.s
 set c = .newc(0,0,0)
 layer = obj.tabaka
 set yaz=.maketext(c,yazi,0,0,obj.sc,obj.angle,"L",layer)
 yaz.p1 = c
 yaz.wsc = obj.wsc
 while .WalkObject("Yapıştırılacak Yeri Seçiniz, [Harita Akademi]",c,8,yaz)
 .addobject(yaz)
 wend

.netcadcommand("REGEN")
 end with
 end sub
