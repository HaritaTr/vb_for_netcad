
function layer_select(list,index)
dim tara,ad
with netcad
index=.numlayers-1
for tara=0 to index
list(tara+1,1)=.layernameof(tara)
next
end with
end function
 
 
 
sub main
dim dialog,list(300,2),tara,index,pagediv,pagecount,page,pagestart,lname,lcod,delcount
layer_select list,index
with netcad
pagediv=10
delcount=0
pagecount=round(index/pagediv,0)
if pagecount=0 then pagecount=1
pagestart=1
for page=1 to pagecount
set dialog = Netcad.NewBDialog("Tabaka Sil ["&pagestart&" - "&pagestart+pagediv&" arasý] [Harita Akademi, Þaban GÜL]")
dialog.PutPrompt " Aþaðýdan Seçilen Tabakalardaki Tüm Objeler Silinecektir."
for tara=pagestart to pagestart+pagediv
if list(tara,1)<>"" then
dialog.GetCheck "item"&tara,list(tara,1),0
end if
next
 
if dialog.showmodal then
for tara=pagestart to pagestart+pagediv
if list(tara,1)<>"" then
list(tara,2)=dialog.ValueByName("item"&tara)
end if
next
end if
pagestart=pagestart+pagediv+1
next
for tara=1 to index
if list(tara,2)=1 then
lname=list(tara,1)
with nclayermanager
lcod=.find(lname)
if lcod<>-1 then
.Delete .find (lname),true
delcount=delcount+1
netcad.netcadcommand("REGEN")
end if
end with
end if
next
if delcount<>0 then msgbox "Toplam "&cstr(delcount)&" Tabaka Silindi",64,"Tabaka Silme"
end with
end sub
 

