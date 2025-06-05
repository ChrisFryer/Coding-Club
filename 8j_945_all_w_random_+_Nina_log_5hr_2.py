import time, os, datetime
import RPi.GPIO as GPIO
import random as rnd
with open('/etc/hostname', 'r') as file:
    nm = file.read().replace('\n', '')
#
def countdown(h, m, s): # Create class that acts as a countdown    
    total_seconds = h * 3600 + m * 60 + s # Calculate the total number of seconds
    while total_seconds > 0:
        timer = datetime.timedelta(seconds = total_seconds)
        print(timer, end="\r")
        time.sleep(1)
        total_seconds -= 1
#
print ("")
print (" > First or outermost treat delivery timer") ; print ("")
#
h = 5
m = rnd.randint(0, 60) #-- randint(start, end)
s = rnd.randint(0, 60)
countdown(int(h), int(m), int(s))
#
curr = datetime.datetime.now()
now = curr.strftime('%Y-%m-%d,%H:%M:%S.%f')[:-4]
f = open("/mnt/nina/new.txt", "a")
f.write(now +"," +nm +",i,dog_fed,1.0\n")
f.flush()
f.close()
#
GPIO.setmode(GPIO.BCM)
GPIO.setwarnings(False)
pin = [2] # GPIO number
GPIO.setup(pin, GPIO.OUT)
GPIO.output(pin, GPIO.HIGH)
try:
  GPIO.output(pin, GPIO.LOW)
  print ("")
  print ("Delivery ON")
except KeyboardInterrupt:
  print ("  Quit")
  GPIO.cleanup()
time.sleep(5);
try:
  GPIO.output(pin, GPIO.HIGH)
  print ("Delivery OFF")
except KeyboardInterrupt:
  print ("  Quit")
  GPIO.cleanup()
print ("")
time.sleep(3);

def countdown(h, m, s): # Create class that acts as a countdown    
    total_seconds = h * 3600 + m * 60 + s # Calculate the total number of seconds
    while total_seconds > 0:
        timer = datetime.timedelta(seconds = total_seconds)
        print(timer, end="\r")
        time.sleep(1)
        total_seconds -= 1
#
print ("")
print (" > Second or innermost treat delivery timer") ; print ("")
#
h = 2
m = rnd.randint(0, 60)
s = rnd.randint(0, 60)
countdown(int(h), int(m), int(s))
#
curr = datetime.datetime.now()
now = curr.strftime('%Y-%m-%d,%H:%M:%S.%f')[:-4]
f = open("/mnt/nina/new.txt", "a")
f.write(now +"," +nm +",i,dog_fed,2.0\n")
f.flush()
f.close()
#
GPIO.setmode(GPIO.BCM)
GPIO.setwarnings(False)
pin = [3] # GPIO number
GPIO.setup(pin, GPIO.OUT)
GPIO.output(pin, GPIO.HIGH)
try:
  GPIO.output(pin, GPIO.LOW)
  print ("")
  print ("Delivery ON")
except KeyboardInterrupt:
  print ("  Quit")
  GPIO.cleanup()
time.sleep(5);
try:
  GPIO.output(pin, GPIO.HIGH)
  print ("Delivery OFF")
except KeyboardInterrupt:
  print ("  Quit")
  GPIO.cleanup()
print ("")
time.sleep(3);
