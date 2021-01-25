' Açıklama : Seçilen alan objelerinin toplam alanını vermeyi sağlayan makro.

Sub Main
Dim i,sayi,secim,obje,alan
  with Netcad
    Set secim=.NewSelectionset
    secim.SELECT "Alan toplamlarını hesaplamak istediğiniz alan objelerini seçiniz.", Array(opline)
    If .EscPressed then Exit Sub
    sayi=secim.NE-1
    alan=0
    Set obje=.NewObject
    For i=0 to sayi
        secim.GetSelectedObject i, obje
        alan=alan + obje.Area
    Next
    .Message 0, "Toplam Alan: " & Round((alan),4) & " m² ", "Alan hesaplandı!"
  end with
End Sub

