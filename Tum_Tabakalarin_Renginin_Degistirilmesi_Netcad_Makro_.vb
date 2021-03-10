
Sub Main
 Dim i
 dim a
 dim bd,ruhan ,elifyaren
 with Netcad





set BD = Netcad.NewBDialog("Tüm Tabakaların Renginin Değiştirilmesi [Harita Akademi, Şaban GÜL]")
 BD.Getfloat "sagulnet","Renk Kodu Giriniz ( Siyah:0 )",0,0
 BD.GETCHECK "sagulnet2","Kilitli Tabakalarda Yapılmasın",1
 if BD.showmodal then
 ruhan=round(BD.ValueByName("sagulnet"),0)
 elifyaren=BD.ValueByName("sagulnet2")
 else
 exit sub
 end if
 for i= 0 to .NumLayers-1

with nclayermanager
 if elifyaren=1 and .layer(i).LockActive=TRUE then
 else
 .Layer(i).color = ruhan

end if

end with

next

.netcadcommand("REGEN")



end with

End Sub
