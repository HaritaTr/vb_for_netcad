
Sub Main
 Dim p, i,ToplamY, ToplamX, ToplamZ
 Dim AgMerY, AgMerX, AgMerZ

with Netcad
 set p = .NewPoly
 if .GetPolygon("Kapalý Alan Seç",p) then



for i=0 to p.num-2
 ToplamY= p.cor(i).y + ToplamY
 ToplamX= p.cor(i).x + ToplamX
 ToplamZ= p.cor(i).z + ToplamZ
 next

AgMerY = Round(ToplamY / (p.num-1),3)
 AgMerX = Round(ToplamX / (p.num-1),3)
 AgMerZ = Round(ToplamZ / (p.num-1),3)

msgbox "Toplam Nokta Sayýsý = " & p.num-1 & chr(13) & "#Aðýrlýk Merkezi Koordinatlarý#" & chr(13) & "Y =" & AgmerY & chr(13) & "X =" & AgMerX & chr(13) & "Z =" &AgMerZ ,64,"Harita Akademi, Þaban GÜL"
 end if

set p = nothing
 end with
 End Sub
