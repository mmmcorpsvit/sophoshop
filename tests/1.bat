"C:\Program Files\GIMP 2\bin\gimp-console-2.8.exe" -i --verbose -b "(crop \"in.jpg\")" -b "(gimp-quit 0)" --verbose
exit

@echo on
cd "C:\Program Files\GIMP 2\bin\"
set gimp="..\plug-ins\crop-auto.exe"
%gimp% "in.jpg"



rem "C:\Dev\sophoshop\_private\ImageMagic\convert.exe" "C:\Dev\sophoshop\tests\images\in.jpg" -trim "C:\Dev\sophoshop\tests\images\in.jpg"
rem "C:\Program Files\GIMP 2\bin\gimp-2.8.exe" -i -b "(autocrop "C:\Dev\sophoshop\tests\images\in.jpg")" -b "(gimp-quit 0)"
rem "C:\Program Files\GIMP 2\bin\gimp-console-2.8.exe" -i -b "(autocrop "C:\Dev\sophoshop\tests\images\in.jpg")" -b "(gimp-quit 0)" --verbose
rem %gimp% -i -b "(flatten-layer-groups 1 in.jpg)" -b "(gimp-quit 0)"



