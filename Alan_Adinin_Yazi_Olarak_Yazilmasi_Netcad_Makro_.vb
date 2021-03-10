SUB Main
 DIM ss,o,i,j,oo,p,sel,poly,tabaka,yazi,a,c ,bd, secenek
 DIM kt() ,t()
 dim elif,ruhan,saban,sagul1,sagul2

With netcad

set BD = Netcad.NewBDialog("Alan Adı Yazdırma [Harita Akademi, Şaban GÜL]")

'BD.Getinteger "item1","Yatay Öteleme Miktarı Giriniz:",0
 'BD.Getinteger "item2","Dikey Öteleme Miktarı Giriniz:",0
 'BD.GetCombo "tabaka1", "İşlem Görecek Tabaka Seçiniz: ", "0", 0
 BD.GetCombo "tabaka2", "Parsel No Hangi Tabakaya Yazılsın: ", "0", 0
 for i = 1 to .numlayers - 1
 BD.AddCombo .LayerNameOf(i)
 next
 BD.Getfloat "yaziboy","Parsel Numaraların Yazı Boyu: ",2,2

if BD.showmodal then

elif= BD.ValueByName("tabaka1")
 ruhan= BD.ValueByName("tabaka2")
 saban= BD.ValueByName("yaziboy")
 sagul1=BD.ValueByName("item1")
 sagul2=BD.ValueByName("item2")



a=.getparam(94)*2/1000
 set SEL = .NewSelectionSet
 set o = .NewObject
 set poly=.newpoly

.setparam beginblock,true
 if SEL.SELECT("Kapalı çoklu doğruları seç",array(opline)) then
 for i = 0 to SEL.NE-1
 j = SEL.GetSelectedObject(i, o)
 set poly=.getplineext(o)
 set c = poly.CenterOfMass



.AddObject (.MakeText (c,o.pname, 0,0, saban,0,"M",.CreateLayer(ruhan,2)))
 next
 .setparam endblock,true
 ' .NetcadCommand("REDRAW")
 set sel = nothing
 set poly = nothing
 set o = nothing
 end if



END if
 End With
 END SUB

