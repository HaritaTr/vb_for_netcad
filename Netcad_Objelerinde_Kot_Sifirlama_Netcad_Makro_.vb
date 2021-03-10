
Sub main
Dim i,j,o,p

dim BD
dim sabangul
dim elifyaren
 set BD = Netcad.NewBDialog("Kot Sýfýrlama [Harita Akademi, Þaban GÜL]")

BD.GetFloat "sagul1","Sabit Kot Deðeri ( Sýfýrlamak için 0 Yazýnýz) ","0" ,3
 BD.Getradio "sagul2","Çoklu Doðrudada Kot Sýfýrlama Yapýlsýnmý ? ","Evet|Hayýr" ,1


 
 if BD.showmodal then
 else
 exit sub
 end if



elifyaren= BD.ValueByName("sagul1")



with Netcad
 for i = 0 to .numobject-1
 set o = .getobject(i)





if o.tag = opline then
 set p = .getplineext(o)
 for j = 0 to p.num-1
 p.cor(j).z = elifyaren
 next
 .putplineext o,p
 .putobject i,o
 set p = nothing
 elseif o.tag = oLine and BD.ValueByName("sagul2")=0 then
 o.p1.z = elifyaren
 o.p2.z = elifyaren
 o.p3.z = elifyaren
 .putobject i,o
 end if
 .drawobject o,14
 set o = nothing
 next
 end with
End Sub
