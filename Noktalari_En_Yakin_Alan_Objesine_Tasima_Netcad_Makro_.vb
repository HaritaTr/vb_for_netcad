
Sub Main
Dim i,BD,obj,POLY,j,lineOBJ,say
Dim BD0
dim saban1,saban2,murat
with Netcad







Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fso, objFolder,f

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFolder = FSO.GetFolder("C:\")
If Not FSO.FolderExists("C:\Sagul") Then
objFolder.SubFolders.Add "Sagul"
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFolder = FSO.GetFolder("C:\Sagul")
If Not FSO.FolderExists("C:\Sagul\Netcad") Then
objFolder.SubFolders.Add "Netcad"
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFolder = FSO.GetFolder("C:\Sagul\Netcad")
If Not FSO.FolderExists("C:\Sagul\Netcad\Makro") Then
objFolder.SubFolders.Add "Makro"
End If




Set fso = CreateObject("Scripting.FileSystemObject")
Set objFolder = FSO.GetFolder("C:\Sagul\Netcad\Makro")
If Not FSO.FolderExists("C:\Sagul\Netcad\Makro\Tanimlar") Then
objFolder.SubFolders.Add "Tanimlar"
End If




Dim fsot, ft
Set fsot = CreateObject("Scripting.FileSystemObject")
Set ft = fsot.OpenTextFile("C:\Sagul\Netcad\Makro\Tanimlar\noktacek.sagul", ForWriting, True)




dim ruhangul

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("C:\Sagul\Netcad\Makro\Tanimlar\ncek-tampon.sagul", ForReading, True)

do while not f.AtEndOfStream
ruhangul= f.ReadLine()
loop







set BD0 = .NewBDialog("Noktalarý en yakýn alan objeye taþýma [Harita Akademi, Þaban GÜL]")
 BD0.GetCheck "SAGUL1", "Alanlar Herhangi Bir Tabakadadýr", -1
 BD0.GetCheck "SAGUL2", "Noktalar Herhangi Bir Tabakadadýr", -1

if BD0.ShowModal then
 saban1=BD0.ValueByName("SAGUL1")
 saban2=BD0.ValueByName("SAGUL2")
 if saban1=0 and saban2=0 then ft.WriteLine 0
 if saban1=0 and saban2=1 then ft.WriteLine 1
 if saban1=1 and saban2=0 then ft.WriteLine 2
 if saban1=1 and saban2=1 then ft.WriteLine 3
 ft.close

if saban1=0 and saban2=0 then murat= 0
 if saban1=0 and saban2=1 then murat= 1
 if saban1=1 and saban2=0 then murat= 2
 if saban1=1 and saban2=1 then murat= 3




else
 exit sub
 end if




set BD = .NewBDialog("Noktalarý en yakýn alan objeye taþýma [Harita Akademi, Þaban GÜL]")

if saban1=0 then
 BD.GetCombo "PARSEL_ALAN", "Alanlarýn bulunduðu tabakayý seçiniz : ", "0", 0
 for i = 1 to .numlayers - 1
 BD.AddCombo .LayerNameOf(i)
 next
 end if




if saban2=0 then
 BD.GetCombo "PARSEL_NOKTA", "Noktalarýn bulunduðu tabakayý seçiniz : ", "0", 0
 for i = 1 to .numlayers - 1
 BD.AddCombo .LayerNameOf(i)
 next
 end if

if ruhangul<0.002 then
 ruhangul=0.01
 end if
 BD.GetFloat "TAMPON", "Tampon Mesafesi (Metre Cinsinden): ",ruhangul, 3
 BD.GetCheck "TAMPON2", "Mesafeyi bir sonraki iþlemler için sakla", 1










' BD.GetCheck "LIMIT_BUL", "Ýþlem Sonrasý Ýþlem Gören Parsellerde Limit Bul", -1
 ' BD.GetCheck "LIMIT", "Ýþlem Sonrasý Tüm Limit Bul.", -1




if BD.ShowModal then




dim sagulnet
sagulnet= BD.ValueByName("TAMPON2")

Dim fsot2, ft2
Set fsot2 = CreateObject("Scripting.FileSystemObject")
Set ft2 = fsot.OpenTextFile("C:\Sagul\Netcad\Makro\Tanimlar\ncek-tampon.sagul", ForWriting, True)
 if sagulnet=1 then
 ft2.WriteLine BD.ValueByName("TAMPON")
 else
 ft2.WriteLine 0
 end if
 ft2.close







set obj = .newobject()
 if murat=0 or murat=1 then

.SetFilter nothing, array(BD.ValueByName("PARSEL_ALAN")), array(opline)
 else
 .SetFilter nothing, array(), array(opline)
 end if
 while .GetNextObject2(obj)
 say = say + 1
 set POLY = .newpoly()
 set POLY = obj.GetObjectAsPline()
 if BD.ValueByName("LIMIT_BUL") = 1 then
 .SetCurrentWindow obj.limits, true
 .DrawObject .MakePline("",3,0,0,0,0,POLY), 222
 end if
 '************************************
 for j = 0 to POLY.Num - 2
 set lineOBJ = .newobject()
 set lineOBJ = .MakeLine(POLY.Cor(j),POLY.Cor(j+1),0,0,3)
 pointMove lineOBJ,BD.ValueByName("PARSEL_NOKTA"),BD.ValueByName("TAMPON")
 next
 '************************************
 set POLY = nothing
 .BackMessage : .setMessage "[ " & say & ". ] Alan Objesi Tarandý"
 wend
 .ResetFilter
 set obj = nothing
 .Message 0, say & " Adet Alan Ýþlem Gördü.", "Belirlenen kritere göre noktalar ayný tabakada taþýndý."
 end if
 .backMessage
 if BD.ValueByName("LIMIT") = 1 then .findworld
end with
End Sub










Function pointMove(lineOBJ,pointLAYER,TAMPON)
dim i,ext,obj,POLY,lenght1,lenght2
dim fso,f ,sabangulX




Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("C:\Sagul\Netcad\Makro\Tanimlar\noktacek.sagul", ForReading, True)

do while not f.AtEndOfStream
sabangulX= f.ReadLine()
loop
 f.close







with netcad
 set ext = .NewWorld(0,0,0,0)
 set ext = lineOBJ.Limits
 ext.Expand TAMPON, TAMPON
 set obj = .newobject()
 if sabangulX="1" or sabangulX="3" then
 .SetFilter ext, array(),array(opoint)
 else
 .SetFilter ext, array(pointLAYER), array(opoint)
 end if
 while .GetNextObject2(obj)
 if NCMath.OnlineSeg(obj.p1,lineOBJ.p1,lineOBJ.p2,TAMPON) then
 lenght1 = NCMath.Distance(obj.p1,lineOBJ.p1,false)
 lenght2 = NCMath.Distance(obj.p1,lineOBJ.p2,false)
 if lenght1 < lenght2 then
 obj.p1.x = lineOBJ.p1.x
 obj.p1.y = lineOBJ.p1.y
 obj.p1.z = lineOBJ.p1.z
 .PutObject .CurObjPos, obj
 else
 obj.p1.x = lineOBJ.p2.x
 obj.p1.y = lineOBJ.p2.y
 obj.p1.z = lineOBJ.p2.z
 .PutObject .CurObjPos, obj
 end if
 end if
 wend
 .resetfilter
 set obj = nothing
end with
end function
