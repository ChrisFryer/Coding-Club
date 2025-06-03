<#

    name: Vinyl-2025.06.03.a0 . ps1
    Created on tiffany , Tested on Rebecca
#>

    $bs = "W:\projects\vinyl records\new\" 
#    $bs = "U:\Master\Desktop\"
    $ErrorActionPreference="SilentlyContinue" #-- set error display to Nil
    function prompt { "> " } ; cls            #-- set prompt to almost nothing
    write-host ''
    $pa = $pc = $pi = "(nil)"
    $srt = 1
    $mg_bk = @{ForegroundColor="magenta";BackgroundColor="black"}  #-- use @bk_gy
    $yw_dg = @{ForegroundColor="Yellow" ;BackgroundColor="DarkGreen"}
    $rd_wh = @{ForegroundColor="red" ;BackgroundColor="white"}
    $wh_gn = @{ForegroundColor="white" ;BackgroundColor="green"}

  do{ 
  cls <#-- clear screen
    output syntax = 3 ,D,Pin,4,5,6,half,8, .txt
    output          nu,a,c,d  ,f,k
    Write-Host ' srt = ' $srt #>
    if ( $srt -lt 2 ){
      write-host '' ;  write-host ' ... ' -NoNewline ; write-host 'Note : ' @mg_bk
      write-host ' ... ' -NoNewline ; write-host 'make sure notepad settings for ' @yw_dg
      write-host ' ... ' -NoNewline ; write-host "'when notepad opens' " @rd_wh
      write-host ' ... ' -NoNewline ; write-host 'is set to ' @yw_dg
      write-host ' ... ' -NoNewline ; write-host "'start newsession & discard unsaved changes' " @rd_wh
      write-host '' ;  $srt += 1 # ; Write-Host ' srt = ' $srt 
      }

#-- Artist
    write-host '' ; write-host '     Artist or Group ' @wh_gn
    write-host ' ... include: ' ; write-host ' ... z_compliation '
    write-host ' ... z_sound track '
    write-host ''
    write-host  ' previous artist : ' $pa @wh_gn  ; $ar = read-host ' '
    if ( $ar.Length -eq 0 ){ 
      if ($pa -eq "(nil)" ){ $ar = ' ' }
        else { $ar = $pa }
        }
     else {
       $ar = $ar.substring(0,1).toupper()+$ar.substring(1) # make the first letter caps
     
       write-host '' }
#-- the quit scenario
    if ( ($ar).ToLower() -eq "q" ){ write-host ' Exiting script ' ; break }
     else { write-host  ' using artist : ' $ar }   

#-- Album
    write-host '' ; write-host '    Album = ?' @wh_gn ; write-host  ' previous album = ' $pc ; $cr = read-host ' '
    if ( $cr.Length -eq 0 ){ 
      if ($pc -eq "(nil)" ){ $cr = ' ' }
        else { $cr = $pc }
        }
    else { write-host '' }
    #-- the quit from here scenario
    if ( ($cr).ToLower() -eq "q" ){ write-host ' Exiting script ' ; break }
    else { write-host  ' using album = ' $cr }

#-- Secondary Album tag
    write-host '' ; write-host '    sub Album Notation = ?' @wh_gn ; $dr = read-host ' '
    if ( $dr.Length -eq 0 ){ $dr = " " }
    else { write-host '' }
    #-- the quit from here scenario
    if ( ($dr).ToLower() -eq "q" ){ write-host ' Exiting script ' ; break }
    else { write-host  ' using sub album note = ' $dr }

    # $br=' '

#-- Live
    write-host '' ; write-host '    Live or Blank = ?' @wh_gn
    write-host '  a = blank'
    write-host '  b = live'
    $er = read-host ' '
    if ( $er.Length -eq 0 ){ $er = "a" }
    else { write-host '' }
    switch( $er ){
      a { $fr = "" ;  break }
      b { $fr = "Live" ; break }
      }
    #-- the quit from here scenario
    if ( ($er).ToLower() -eq "q" ){ write-host ' Exiting script ' ; break }
    else { write-host  ' using : ' $fr }

#-- Type
    write-host '' ; write-host '    Attributes or Type : ' @wh_gn
    write-host '  a = blank'
    write-host '  b = double'
    write-host '  c = Remastered'
    write-host '  d = sterio'y

    $jr = read-host ' '
    if ( $jr.Length -eq 0 ){ $jr = " " }
    else { write-host '' }
    switch( $jr ){
      a { $kr = " " ;  break }
      b { $kr = "double" ;  break }
      c { $kr = "Remastered" ; break }
      d { $kr = "sterio" ;  break }
      }
    #-- the quit from here scenario
    if ( ($kr).ToLower() -eq "q" ){ write-host ' Exiting script ' ; break }
    else { write-host  ' using : ' $kr }

#-- 33/45
    write-host '' ; write-host '   record type = ?'  @wh_gn
    # ; $gr = read-host ' '
    
    write-host " a = 33 single EP"
    write-host " b = double"
    write-host " c = 45 single"
    write-host " d = sterio"
    write-host " e = wanted"
    write-host " f = else"
    write-host""   
    $hr = read-host ' '
    if ( $hr.Length -eq 0 ){ $gr = " " ; $nu = "a" }
#    else { write-host 'using :' $nu}

    switch($hr){
      "a"{ $hr = "" ;        $nu="3" ; break }
      "b"{ $hr = "double" ;  $nu="3" ; break }
      "c"{ $hr = "single" ;  $nu="4" ; break }
      "d"{ $hr = "sterio" ;  $nu="3" ; break }
      "e"{ $hr = "wanted" ;  $nu="5" ; break }
      "f"{ $hr = ""       ;  $nu="6" ; break } 
      }
    #-- the quit from here scenario
    if ( ($hr).ToLower() -eq "q" ){ write-host ' Exiting script ' ; break }
    else { write-host  ' using : ' $hr }

#-- Notes
    write-host '' ; write-host '    Notes :' @wh_gn ; write-host  ' previous notes = ' $pi 
    $ir = read-host ' '
    if ( $ir.Length -eq 0 ){ 
      if ($pi -eq "(nil)" ){ $ir = ' ' }
        else { $ir = $pi }
        }
    else { write-host '' }
    #-- the quit from here scenario
    if ( ($ir).ToLower() -eq "q" ){ write-host ' Exiting script ' ; break }
    else { write-host  ' using notation : ' $ir ; write-host ''}
   
    write-host ' ... Write file with Name info set as : '@wh_gn ; write-host ''
#-- Check / Confirm Name
    $nm= $nu+","+$ar+","+$cr+","+$dr+","+$fr+","+$kr+","+$hr+","+$ir+","
    write-host $nm        
    $pa = $ar ; $pc = $cr ; $pi = $ir #-- Make this previous iterations persist
    start-sleep 1

    write-host '' ; write-host ' ... is this name OK ? ' @mg_bk ; write-host ''
    write-host "     " -NoNewline ; write-host "'n' or 'no'  will restart this script without writing out " @rd_wh
    write-host ''
    $mr = read-host " default action is enter  for 'Yes' "
    if ( ( ($mr).tolower() = "no" ) -or ( ($mr).tolower() = "n" ) ){
      write-host ' no requstered '
      }
    else{

#-- write out file name
    $ap = "notepad" ;  start-process $ap ; $w = "WASP" <#
    I am using an early Windows Automation Snapin for PowerShell (WASP) 
    portable version about 1.2* to control notepad from Powershell #>
    if ( ( Get-Module -Name $w ).name -ne $w ) { Import-Module U:\_ps\WASP\WASP.dll } ; start-sleep 1
    Select-Window $ap | Set-WindowActive | Send-Keys "^{s}" ;  start-sleep 1 #-- save file
    Select-Window $ap | Select-ChildWindow | Set-WindowActive | Send-Keys $bs ; Start-Sleep 1 
    Select-Window $ap | Select-ChildWindow | Set-WindowActive | Send-Keys $nm ; Start-Sleep 1 
    Select-Window $ap | Select-ChildWindow | Set-WindowActive | Send-Keys ".txt" ; Start-Sleep 1 
    Select-Window $ap | Select-ChildWindow | Set-WindowActive | Send-Keys `t`t`t`t ; start-sleep 1
    Select-Window $ap | Select-ChildWindow | Set-WindowActive | Send-Keys "{enter}"
    Select-Window $ap | Set-WindowActive | Send-Keys "%{f}" ; start-sleep 1
#    Select-Window $ap | Set-WindowActive | Send-Keys "%{}" ; start-sleep 1
    Select-Window $ap | Set-WindowActive | Send-Keys "%{x}" ; start-sleep 1
#    Select-Window $ap | Set-WindowActive | Send-Keys "x"

    } #-- end confirmation else

  } until ( ($ar).tolower() -eq "q")

  write-host '' ; Write-Host ' ... script finished .. '
  Write-Host ' ... Press F5 to restart this script '

#-- ende