
function main
 dim file,fso,str,arr,excSheet,excLine
 dim i
 ' set fso = CreateObject("Scripting.FileSystemObject")
 ' set file = fso.openTextFile(path,1)
 set excSheet = CreateObject("Excel.Application")
 excSheet.workbooks.add
 excLine = 2
 excSheet.cells(1,1) ="Harita Akademi Netcad Makrolarý Ýle Üretilmiþtir."
 excSheet.cells(2,1) ="Tabaka Adý"
 excSheet.cells(2,2) ="Tabaka Rengi"
 excSheet.cells(2,3) ="Tabaka Rengi RGB"
 excSheet.cells(2,4) ="Tabaka Açýk"
 excSheet.cells(2,5) ="Tabaka Görünür"
 excSheet.cells(2,6) ="Tabaka Kilit"
 excSheet.cells(2,7) ="Tabaka Yazýcý Gönder"
 excSheet.cells(2,8) ="Tabaka Baþ. Ölçeði"
 excSheet.cells(2,9) ="Tabaka Bit. Ölçeði"

with ncLayermanager
 for i = 0 to .numLayer-1
 excLine = excLine + 1
 excSheet.cells(excLine,1)= .layer(i).Name
 excSheet.cells(excLine,2)= .layer(i).color
 excSheet.cells(excLine,3)= .layer(i).colorBGR
 excSheet.cells(excLine,4)= .layer(i).IsOpen
 excSheet.cells(excLine,5)= .layer(i).VisibiltyActive
 excSheet.cells(excLine,6)= .layer(i).LockActive
 excSheet.cells(excLine,7)= .layer(i).PrintableActive
 excSheet.cells(excLine,8)= .layer(i).VisStartScale
 excSheet.cells(excLine,9)= .layer(i).VisEndScale

next
 excSheet.cells(excLine+1,1) =".........................................................................."
 excSheet.cells(excLine+2,1) ="Harita Akademi, Þaban GÜL"
 excSheet.cells(excLine+3,1) ="Daha fazlasý için haritaakademi@gmail.com"
 end with
 excSheet.visible = True
 set fso = nothing
 set file = nothing
 end function
