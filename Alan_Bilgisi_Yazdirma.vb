SUB Main
DIM ss,o,i,j,oo,p,sel,poly,tabaka,yazi,a,c   ,SAGUL_DLG, secenek
DIM kt() ,t()
dim elif,ruhan,saban,saban_olcek
 dim saban1,saban2,saban3,saban4
 dim elif1,elif2,elif3,sagul67

With netcad
  saban_olcek=.getparam(94)
with nclayermanager
.add 0,4
end with  

set SAGUL_DLG = Netcad.NewBDialog("Alan Objelerinde Bilgi Yazdýrma , [Þaban GÜL, Harita Akademi]")

      SAGUL_DLG.GetRadio "elif", "Hangi Bilgiyi Yazdýrayým ?", "Hesap Alaný|Tapu Alaný|Gýs Sýnýfý|Gýs Adý|Tabakasý|Alan Adý", 0
      SAGUL_DLG.GetCombo "yaren", "Hangi Tabakaya Yazayým ?", 0,0
           for i = 1 to .numlayers - 1
         SAGUL_DLG.AddCombo .LayerNameOf(i)
     next
      SAGUL_DLG.Getfloat "yaziboy","Yazý boyu kaç olsun ? ",2,2
      SAGUL_DLG.GetCheck "sagulnet1", "Yazý Ýtalik Olsun", 0
      SAGUL_DLG.GetCheck "sagulnet2", "Yazý Altçizgili Olsun", 0
      SAGUL_DLG.GetCheck "sagulnet3", "Yazý Arka Fon Olsun", 0




   if SAGUL_DLG.showmodal then

     saban1=    SAGUL_DLG.ValueByName("elif")   'tür
     saban2=    SAGUL_DLG.ValueByName("yaren")  'tabaka
     saban3=    SAGUL_DLG.ValueByName("yaziboy") 'boyu


     elif1=  SAGUL_DLG.ValueByName("sagulnet1")
     elif2=  SAGUL_DLG.ValueByName("sagulnet2")
     elif3=  SAGUL_DLG.ValueByName("sagulnet3")






      if elif1+elif2+elif3=0 then sagul67=0

                 set SEL = .NewSelectionSet
                 set o = .NewObject
                 set poly=.newpoly

                .setparam beginblock,true
                if SEL.SELECT("Kapalý çoklu doðrularý seç, Þaban GÜL, Harita Akademi",array(opline)) then
                 for i = 0 to SEL.NE-1
                 saban3=saban3/1000*saban_olcek
                   j = SEL.GetSelectedObject(i, o)
                   set poly=.getplineext(o)
                    set c = poly.CenterOfMass
                    if saban1=5 then saban4=o.pname
                    if saban1=0 then saban4= round(o.area,3)
                    if saban1=1 then saban4=round(o.tarea,3)
                    if saban1=2 then saban4=o.cls
                    if saban1=3 then saban4=o.objname
                    if saban1=4 then
                    saban4=o.tabaka
                     with nclayermanager
                    saban4= .layer(saban4).name
                    end with

                    end if
                    if saban1=6 then saban4=o.renk

                   '.AddObject (.MakeText (c,saban4, sagul67,0, saban3,0,"M",.CreateLayer(saban2,2)))
      if   elif1=0 and  elif2=0 and elif3=0 then .AddObject (.MakeText (c,saban4, 0,0, saban3,0,"M",.CreateLayer(saban2,2)))
      if   elif1=1 and  elif2=0 and elif3=0 then .AddObject (.MakeText (c,saban4, 1,0, saban3,0,"M",.CreateLayer(saban2,2)))
      if   elif1=0 and  elif2=1 and elif3=0 then .AddObject (.MakeText (c,saban4, 2,0, saban3,0,"M",.CreateLayer(saban2,2)))
      if   elif1=0 and  elif2=0 and elif3=1 then .AddObject (.MakeText (c,saban4, 4,0, saban3,0,"M",.CreateLayer(saban2,2)))
      if   elif1=1 and  elif2=1 and elif3=0 then .AddObject (.MakeText (c,saban4, 1+2,0, saban3,0,"M",.CreateLayer(saban2,2)))
      if   elif1=1 and  elif2=0 and elif3=1 then .AddObject (.MakeText (c,saban4, 1+4,0, saban3,0,"M",.CreateLayer(saban2,2)))
      if   elif1=0 and  elif2=1 and elif3=1 then .AddObject (.MakeText (c,saban4, 2+4,0, saban3,0,"M",.CreateLayer(saban2,2)))
      if   elif1=1 and  elif2=1 and elif3=1 then .AddObject (.MakeText (c,saban4, 1+2+4,0, saban3,0,"M",.CreateLayer(saban2,2)))


                 next
                 .setparam endblock,true
                 .NetcadCommand("REDRAW")
                 set sel = nothing
                 set poly = nothing
                 set o = nothing
             end if


  END if
End With
END SUB
