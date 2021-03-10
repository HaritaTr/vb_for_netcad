
Sub Main
Dim Top
Dim obj
dim bd
dim bz
dim result
 with Netcad
 set BD = Netcad.NewBDialog("Nokta Adı Değiştirme [Harita Akademi, Şaban GÜL]")

BD.PutPrompt "Aşağıdaki Alana Eski Karakteri Giriniz:"
 BD.Getstring "sagul1","Eski: ","/",50

BD.PutPrompt "Yerine Yazılacak Yeni Karakteri Giriniz:"
 BD.Getstring "sagul2","Yeni: ","_",50
 BD.PutPrompt " "





if BD.showmodal then



if BD.ValueByName("sagul1")="" then
 result = MsgBox ("Eski Karakter Kısmı Boş Bırakıldı Devam etmek istermisiniz ", vbYesNo, "[Harita Akademi, Şaban GÜL]")
 Select Case result
 Case vbYes
 Case vbNo
 exit sub
 End Select
 end if



if BD.ValueByName("sagul2")="" then
 result = MsgBox ("Yeni Karakter Kısmı Boş Bırakıldı Devam etmek istermisiniz ", vbYesNo, "[Harita Akademi, Şaban GÜL]")
 Select Case result
 Case vbYes
 Case vbNo
 exit sub
 End Select
 end if



.setfilter nothing, array(),array(opoint)
 do
 set obj=.getnextobject

if obj is nothing then
 exit do
 else
 obj.s=replace(obj.s,BD.ValueByName("sagul1"),BD.ValueByName("sagul2"))
 end if
 .PUTOBJECT .CUROBJPOS,OBJ
 loop
 .resetfilter



end if

result=Msgbox("İşlem Başarıyla Tamamlandı",vbInformation, "[Harita Akademi, Şaban GÜL]")
 end with
End Sub
