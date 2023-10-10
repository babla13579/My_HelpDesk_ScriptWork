 $numberDone = 0
 $log = 'E:\Official Records-TIF\log_written.txt'
 $A = gci -Recurse -File -Path 'E:\Official Records-TIF' -Include *.pdf
 $ACount = $A.Count 

 $A | % {
   $numberDone = $numberDone + 1
   $PDFFullFilePath = $_.Directory,"\",$_.Name -join ""
   $TIFFullFilePath = $PDFFullFilePath -replace 'pdf','tif'

   if(-not(Test-path $TIFFullFilePath -PathType leaf)) {
      & "C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe" convert -density 200 -compress lzw -units 2 $PDFFullFilePath $TIFFullFilePath
      Write-Host $numberDone / $ACount  --- Converting to TIF ---- $PDFFullFilePath
      echo $TIFFullFilePath >> $log
   } else {
     Write-host Progress $numberDone / $ACount  --- TIF ALREADY EXISTS ---  $TIFFullFilePath
   }
 }

