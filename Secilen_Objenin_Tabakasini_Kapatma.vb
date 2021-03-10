Sub Main()
Dim obje
Dim secim
Dim sagulnet1
Dim sagulnet2
Dim sgl_1,sgl_2,sgl_3,sgl_4,sgl_5,sgl_6
Dim a,b,c,d,e,f,g,h
Dim icindeki
Dim tbsagul,tbex
Dim SEL,o,i,j
Dim tabaka, tabaka1, tabaka2, tbksay, toptbk
Dim tabakasayi, aciksayi, kapalisayi, kilitlisayi, kilitsizsayi, yazicigonder, yazicigonderme
Dim list, list1, list2, list3, list4
Dim max,min,elif ,sabangul
dim xls, xlspath,alan,DEG,CL, bd, U,V,R,W,ruhangul, elifyaren
dim saban,ruhan
dim ss

 
with netcad


 set ss = .NewSelectStatus ' Anlik Secim objesi yarat
 while .SelectObjectInstant("Seçtiğiniz Objenin Tabakası Kapatılacaktır.",1,array(),ss)
 set o = ss.objects(0) ' Secim objesinin ilk objesini al
 saban= o.tabaka

With nclayermanager

netcad.closelayer(saban)

End With



 set o = nothing
 wend
 set ss = nothing
   .setcurrentlayer(0)
 end with

MsgBox "İşlem Başarıyla Tamamlandı", 0, "Şaban GÜL, Harita Mühendisi"
end SUB
