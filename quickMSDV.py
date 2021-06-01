#Written by Roger Sheu, last revised 9/25/2019
#
#This program uses pyautogui, so you'll need to install that.
#If you're using Anaconda, go into your Anaconda terminal and type in
#"pip install pyautogui". This should install the package. Any accompanying
#package may be installed at your discretion, too.
#####
import pyautogui
import win32com.client as w32
import time

retry = 3
shell = w32.Dispatch("WScript.Shell")

#To start, select an output folder and a mass list in MSDataView.
#Watch out where you put your output files. If they're in the same folder
#as the data, determine how that will affect the code.

#I have mass lists for CH, CHO, and CHO2. If you want them, just e-mail me.

#Also, make sure there's no path selected next to Select Data when you start!

#Set your command line directory to wherever you've placed this code.
#Use "cd C:\Users\Roger\Dropbox\Research\Scripts..." to direct Python to the right place.

#When ready, type in "python quickMSDV.py" into the prompt and follow
#instructions from there.

#WARNING: This program will OVERWRITE your data if you already have results
#in the destination folder.

print("Preload your mass list and destination folder in MSDataView.")
print("Check that 'Select Data' is highlighted, with no current path.")

#Just click 'Select Data' then cancel if this isn't currently the case.

print("How many samples do you have?")
numSamples = int(input(">"))


#Change "input" to "raw_input" if you're running Python 2.
input("Press Enter key to continue. You will have a few seconds.")
print("Thank you. Make sure you have the MSDataView window selected.")

time.sleep(5)

###Click over to the MSDataView program now, if you haven't already!

for i in range(0,numSamples):
    #Opens file directory menu
    pyautogui.press('enter')
    time.sleep(1)

    #########
    #The following just routes the selection to my own APCI data in the drive.
    #Change the numbers for your own purposes.
    #At the moment, it's going down 10 times, then picking that subfolder,
    #going down 27 times, picking THAT subfolder, then going down another 19
    #times.
    #Then, it goes down as many times as needed
    #(variable for the number of samples).
    #####This means you just have to change the numbers in "range(0,x)"
    #####If you have more or fewer subfolders, duplicate or remove as needed.
    #####
    
    #As such, smart folder arrangement may be beneficial.
    for j in range(0,12):
        pyautogui.press('down')
        time.sleep(0.05)
    pyautogui.press('right')
    time.sleep(2)
    
    for k in range(0,32):#numDown 2:
        pyautogui.press('down')
        time.sleep(0.05)
    pyautogui.press('right')
    time.sleep(2)
    
    for m in range(0,5):#numDown 2:
        pyautogui.press('down')
        time.sleep(0.05)
    pyautogui.press('right')
    time.sleep(2)

    for m in range(0,6):#numDown 2:
        pyautogui.press('down')
        time.sleep(0.05)
    pyautogui.press('right')
    time.sleep(2)


    #for n in range(0,2):#numDown 2:
    #    pyautogui.press('down')
    #pyautogui.press('right')
    #time.sleep(1)
    
    #This check variable was because I had a couple of irrelevant standard runs.
    #The number tells it how many files to skip (if any).
    #Set it to 0 if you want it to go through each file. If any errors occur
    #during run, this will come in handy, since you can start in the middle.
    check=i+3

    pyautogui.press('down')    
    for l in range(0,check):
        pyautogui.press('down')
        time.sleep(0.05)
    pyautogui.press('enter')

    time.sleep(1)
    #Directs the cursor to the Run button
    for j in range(0,4):
        pyautogui.press('tab')
    #Presses run
    pyautogui.press('enter')
    #print("Sample " + str(i+1) + " in progress.")
    for z in range(0,3): 
         print(str(20*(3-z))+ " seconds remaining.")
         time.sleep(20) #Wait for Steven's code to run. Can be adjusted as needed
    #time.sleep(60) 
    print("Sample " + str(i+1) + " complete.")

    ###Reset back to original state.
    pyautogui.keyDown('shift')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.keyUp('shift')
    time.sleep(0.1)
    pyautogui.press('enter') #This presses Clear to the right of "Select Data"
    pyautogui.keyDown('shift')
    pyautogui.press('tab')
    pyautogui.keyUp('shift') #Resets back to original state.
    time.sleep(0.1)
