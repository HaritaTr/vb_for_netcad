
Sub Main
 Dim obj,BD,BD2,sagul,i,RUHAN,j,o,SS,SEL
 with Netcad

set BD = Netcad.NewBDialog("Tapu Alanlarının Hesap Alanı Yapılması [Harita Akademi, Şaban GÜL]")



BD.Getradio "sagulnet","Bir Yöntem Seçiniz","Tüm Projedeki Alan Objeleri|Bir Tabakadaki Alan Objeleri|Ekrandan Tek Tek Seçilen Alan Objeleri| Seçim Kümesi Oluşturarak Alan Seç" ,0



if BD.showmodal then
 sagul= BD.ValueByName("sagulnet")
 if sagul=1 then
 set BD2 = Netcad.NewBDialog("Tapu Alanlarının Hesap Alanı Yapılması [Harita Akademi, Şaban GÜL]")
 BD2.GetCombo "tabaka", "Alanların bulunduğu tabakayı seçiniz : ", 0, 0
 for i = 1 to .numlayers - 1
 BD2.AddCombo .LayerNameOf(i)
 next

if BD2.showmodal then
 else
 exit sub
 end if

if sagul=1 then
 RUHAN= BD2.ValueByName("tabaka")
 end if
 with nclayermanager
 ruhan= .layer(ruhan).name
 end with
 end if

end if

if sagul=2 then
 set ss = .NewSelectStatus ' Anlik Secim objesi yarat
 while .SelectObjectInstant("Tapu Alanı Hesap Alanı Yapılacak Alanları Seç",1,array(oPline),ss)
 set o = ss.objects(0) ' Secim objesinin ilk objesini al
 o.tarea=o.area ' rengini sari yap
 .PutObject ss.indexs(0), o ' objeyi geri koy
 .DrawObject o,-1 ' kendi rengi ile ciz
 set o = nothing
 wend
 set ss = nothing

exit sub
 end if



if sagul=3 then
 with Netcad
 set SEL = .NewSelectionSet ' Yeni kume yarat
 set o = .NewObject
 if SEL.SELECT("Tapu Alanı Hesap Alanı Yapılacak Alan Kümesini Seç",array(opline)) then ' istenen turleri kumeye ekle
 for i = 0 to SEL.NE-1 ' kumenin her bir elemani icin
 j = SEL.GetSelectedObject(i, o) ' objeyi ve gercek indeksini al
 o.tarea = o.area ' rengini sari yap
 .putobject j, o ' objeyi geri koy
 next
 SEL.RedrawAndRewind ' secim kumesini toplu kendi renginde
 end if ' cizdir ve kumeyi basa sardir.
 set SEL = nothing
 set o = nothing
 end with
 exit sub
 end if



if sagul=0 then
 .SetFilter nothing, array(), array(opline)
 end if

if sagul=1 then
 .SetFilter nothing, array(BD2.ValueByName("tabaka")), array(opline)
 end if



do

set obj=.getnextobject
 if obj is nothing then
 exit do
 end if
 .drawobject obj,102
 obj.tarea=o.area
 .PUTOBJECT .CUROBJPOS,OBJ

loop
 end with

End Sub
 
