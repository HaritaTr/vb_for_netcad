Sub Main

with netcad
           
dim pc
set pc = Netcad.NewProjection
''Dikkat : Datum listesine aşağıdaki adresten ulaşabilirsiniz.
           '  http://sabangul.com.tr/nc-pro-datum

'' Projeksiyon Ayarı
'' Burada projeksiyon için bir kod giriniz
'' Örnek: UTM 3 için 3 giriniz. UTM 6 için 2 giriniz. Coğrafi için 1 giriniz.
 pc.ProjectionType =3
'' Datum Ayarı
'' Burada datum için bir kod giriniz
'' Örnek: ITRF için 1, ED50 için 4 , WGS için 0 giriniz.
pc.Datum = 1
 '' Dilim Ayarı
'' Burada dilim için bir kod giriniz
'' Örnek: Ankara için 33 girin, Siirt için 42 girin.
 pc.Zone = 42


 pc.SetToCurrentProject false     ' projenin dönüşmesini engellemek için konuldu



 end with
 msgbox " Projeksiyon Başarıyla Tanımlandı ",0,"Dahası: www.sabangul.com.tr"
 end sub
