MSTeams AutoMembersSweep

  This script automates the process of searching for and removing names from a system by simulating mouse and keyboard actions. It reads the names from an Excel file, iterates through them, and performs the necessary actions to remove each name. The script includes safety checks to exit if the 'q' key is pressed.

Rules for Using the Script
  1. Use with Care:
  
  Ensure that the launch flag is set to False during testing to prevent unintended actions. Only set it to True when you are confident that the script is functioning correctly and you are ready to perform the actual removal operations.
  
  2. Verify Coordinates:
  
  Double-check the screen coordinates for the search bar and remove button (search_coords and remove_coords). Incorrect coordinates can lead to unintended actions, such as clicking on the wrong elements or performing actions outside the intended application.
  
  3. Backup Data:
  
  Always create a backup of your data before running the script. This includes the Excel file containing the names and any system data that will be affected by the script. This precaution ensures that you can restore your data in case of any errors or unintended actions.

Set Up

  Variable Details:
    starting_row: The row number in the Excel file to start reading from (1-based index).
    search_coords: The screen coordinates of the search bar.
    remove_coords: The screen coordinates of the remove button.
    launch: A flag to enable or disable the actual clicking actions (set to False for testing).
    file_path: The path to the Excel file containing the names.
    loadtime: The time to wait for the names to load, in seconds.

Limitation

Special Characters
  For some reason, MS Teams has issued searching for names with the character "ñ" and maybe other characters as well. So be aware and check names manually if necessary.

Disclaimers for Using the Script

  1. Use at Your Own Risk:
  
  This script is provided "as is" without any warranties or guarantees. The user assumes all responsibility for any actions taken     using this script. Ensure you understand the script's functionality and test it thoroughly before using it in a live environment.
  
  2. Verify Before Execution:
  
  Double-check all parameters, including file paths, screen coordinates, and the launch flag, before running the script. Incorrect settings can lead to unintended actions, and the author is not responsible for any errors or data loss resulting from the use of this script.
  
  3. Backup Your Data:
  
  Always create a backup of your data before running the script. This includes the Excel file containing the names and any system data that will be affected by the script. The author is not liable for any data loss or corruption that may occur as a result of using this script.
