As I emailed Chris,
I have some bugs in this:
   1. Alarms not coping with single digits
      Single numerals for minutes and seconds not catered for ATM due to the way time is handled
   2. Alarms with a variation crossing the 60 second barrier and still functioning
      60 being a cross over number is not dealt with well thus far due to the way time is handled
   3. 

Work around:
   To get around the above I have set the time default time for muntes and seconds to 10
   I have removed the vaiation or time deviation / adjustment to stop the scripted creap across that 60 sec threash-hold.

Working:
  it will only trigger each alarm once.
  even if the script is run after the first alarm, that forst aarm will not trigger a second time
  Runs 2 alarms with the option to expand further
