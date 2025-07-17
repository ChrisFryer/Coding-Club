#!/usr/bin/env python3
import subprocess, sys, time, os
from datetime import datetime as dt
#
#-- FIRST  alarm aka : a*2
ah2=14 ; am2=51 ; as2=10    #-- hr ; min ; sec
#-- SECOND alarm aka : a*3
ah3=15 ; am3=10 ; as3=10    #-- hr ; min ; sec
#
af2=str(ah2)+str(am2)+str(as2) ; af3=str(ah3)+str(am3)+str(as3) ; print('')
k2=0 ; k3=0 ; k4=0  #-- ensure actions only trigger once including retrospectively # while True:
   now=dt.now() ; nh=now.strftime("%H%M%S")  #-- now #
   if k2==0:     #-- solely focus on Alarm 1
     df2=int(af2)-int(nh) ; print('   alarm 1 :',af2,'- now :', nh,'= 
diff 1 :',df2 )
     if ( ( df2 < 0 ) and ( df2 > -110 ) ): #-- #110 == 70 seconds as HHMMSS (000110)
       print(' now=', nh,': k2 =', k2,': alarm ONE triggered ... ')
#-- add alarm scripted actions or code to execute
       k2=1 ; print(' now=', nh,': k2 =', k2 )
     elif df2 < 0: k2=1 #-- this catches a time in the past & does not execute the alarm code
     else: pass
#
   elif k3==0:   #-- Then only focus on Alarm 2
     df3=int(af3)-int(nh) ; print('   alarm 2 :',af3,'- now :' , nh,'= 
diff 2 :',df3 )
     if ( ( df3 < 0 ) and ( df3 > -110 ) ):
       print(' now=', nh,': k3 =', k3,': alarm TWO triggered ... ')
#-- add 2nd alarm scripted actions or code to be executed
       k3=1
     elif df3 < 0: k3=1 #-- this catches a time in the past & does not execute the alarm code
     else: pass
#
   else: #-- After Alarm 2 -Display exit instructions & do nothing
     if k4==0:
       print(' now=', nh,': doing nothing else today ' )
       print('      but still checking for alarm 1 trigger tomorrow ...')
       print('      ctl C to exit the infinite loop \n')
       k4=1
     else:
       pass #-- do nothing
#
   time.sleep(20)
