
Sub Main
 Dim BD,by ,bz
 set BD = Netcad.NewBDialog("Nokta,Doðru ve Alanlarda Koordinat Yuvarlama [Harita Akademi, Þaban GÜL]")
 BD.GetInteger "item1","Duyarlýk Giriniz: ",2
 BD.GetFloat "item2","Verilen kot deðerine getir (-1 Deðiþtirme) ",-1,3
 if BD.showmodal then
 by=BD.ValueByName("item1")
 bz=BD.ValueByName("item2")
 end if
 set BD = Nothing
 Dim i,j,o,SEL ,ad,ii,j3
 dim c1 ,x,y,y1,x1,p ,yn,xn



with netcad
 set SEL = .NewSelectionSet ' Seçim kümesi oluþtur.
 set o = .NewObject
 if SEL.SELECT("Duyarlýk deðiþecek objeleri seçiniz...",array(opoint,oline,opline)) then ' istenen turleri kumeye ekle
 for i = 0 to SEL.NE-1
 j = SEL.GetSelectedObject(i, o)
 'Noktalar içim
 if o.tag = 1 then
 yuvarla o.p1.y,o.p1.x,y,x,by
 o.p1.y=y
 o.p1.x=x
 end if
 'Hatlar Ýçin
 if o.tag = 2 and by<>-1 then
 yuvarla o.p1.y,o.p1.x,y,x,by
 o.p1.y=y
 o.p1.x=x
 yuvarla o.p2.y,o.p2.x,y,x,by
 o.p2.y=y
 o.p2.x=x
 end if
 'Çoklu doðrular için
 if o.tag = 7 and by<>-1 then
 set p = .getplineext(o)
 for j3 = 0 to p.num-1
 y1= p.cor(j3).y
 x1= p.cor(j3).x
 yuvarla y1,x1,y,x,by
 p.cor(j3).y=y
 p.cor(j3).x=x
 if bz<>-1 then
 p.cor(j3).z=bz 'isteðe baðlý
 end if
 next
 .putplineext o,p
 end if
 o.renk = yellow
 .PutObject J, o
 NEXT
 SEL.RedrawAndRewind
 set o = nothing
 END if
 set SEL = nothing
 set o = nothing
 end with
 end sub





sub yuvarla(ys,xs,c,d,k)
 c= formatnumber(ys,k ,,,0)
 d= formatnumber(xs,k ,,,0)
 end sub
