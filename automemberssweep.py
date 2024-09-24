"""
 * Program Name: AutoMembersSweep for MSTeams
 * Author: Renz Faminiano
 * Date Created: 9/19/2024
 * Last Modified: 9/23/2024
 * Version: 1.2
 * Description: Designed to streamline the process of removing numerous 
                members from a Microsoft Teams group.
 * Usage: 
        1. Change File Path
        2. In Excel, put all data vertically in Column A with the header
        3. Set MoveTo coordinates, search_coords to the Search Bar and remove_coords to the remove(x) button
        4. Set launch to True
        5. Run
        6. Press (Q) to Exit while running (You may need to click multiple times)
        7. Change starting_row to last run when resuming 
 * Dependencies: 
        pip install pyautogui #0.9.54
        pip install keyboard #0.13.5
        pip install pandas #2.2.2

"""

import pyautogui as pg
import keyboard
import time
import pandas as pd   

#Set up
#Start with row (starts with 1 and doesn't include header)
starting_row = 1 #Change what row to start
search_coords = [1830,120]
remove_coords = [1870, 209]
launch = False #Change to False if on Testing
file_path = 'C:\Users\[YourUsername]\Documents'
loadtime = 1 #Pause time in seconds, modify depending on how fast the names load

#Reads all rows in column A
data = pd.read_excel(file_path)  
names = data.iloc[:,0].tolist()
names = [x for x in names if str(x) != 'nan']

#Main Script

n = starting_row - 1
names = names[n:]

head = names[:5] if len(names) > 0 else []
tail = names[-5:] if len(names) > 0 else []
print(f"Removing from Head: {head} \nto\nTail: {tail}\nRemoving {len(names)} members. ")

def check_exit():
    if keyboard.is_pressed('q'):
        print("Exiting...")
        exit()

for i in names:
    
    n += 1

    check_exit()

    #Move to Search Bar and write the current name
    pg.moveTo(search_coords)
    if launch:
        pg.click() 
        pg.hotkey('ctrl', 'a')
        pg.write(i)
    time.sleep(loadtime)

    check_exit()

    #Move to Remove Button and Click
    pg.moveTo(remove_coords)
    time.sleep(loadtime/10)
    if launch:
        pg.click()
        print(i,"Removed", n)

    check_exit()


