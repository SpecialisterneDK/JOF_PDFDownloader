﻿# PDF Downloader
PDF Downloader console program.

## What is?
This PDFDownloader, is a program that read through an excel file with the GRI_2017_2020 format. Do NOT use other formats!
It will download the files provided by the links in the collons AL and AM. Assuming the links are valid.


## How to use?
**1**: Start the program
**2**: Read the prompt on the screen.
**3**: Copy the path to the GRI file and paste it into the program window and click enter.
**4**: Wait until the program Writes "Download Done!" and tells you that you can safely stop the program.
**5**: Check the folder with the GRI file. Profit.

If you have any question regarding the steps, please refer to the "advanced how to" down below.


### Advanced how to
**1: starting the program?**
- Download the program
- If on Github, click on 'Code' in the green box and click 'Download Zip'.
- If you download as zip, put the downloaded file in a folder alone, then right click and press 'unzip here'
- After downloading the program, double click on the "PDFDownloader.sln".
- This will open the program in visual studio
- You will now see a green triangle near the upper part of the program; click it to run the program.

ALTERNATIVE:
- If you have downloaded the program as a .exe, run that instead.(simply double click on it)

**2: Read the prompt on the screen?**
- After starting the program, assuming no errors occur, you should now see a black screen show up with text on it.
- Simply read the text before continuing.

**3: Copy the path?**
- Find the folder that contains the GRI_2017_2020 excel file, open it.
- Some place on the screen you should see names with arrows pointing to the right, click ONCE on it, it should now be marked blue.
- Right after, press the keys 'ctrl' and 'c' at the same time.
- Now back to the program, click on the program window, anywhere on the black screen, and press the 'crtl' and 'v' keys at the same time.
- You should now see new text pop into the window(example below), and if you do, simply press 'Enter'.

Example of a path:
```
 C:\Users\xyz\Desktop\PDFAssignment
```

**4: Wait until done?**
- Yes, just wait until the program writes "Download Done!" and "You can safely close the program".
- Now simply close the program.

**5: Check the folder?**
- In the same folder that you have the GRI excel, you should now see a new folder named "DownloadFolder"
- Inside the new folder, you will find all the pdf's as well as a .txt file.
- The text file contains a report on what pdf's where downloaded, and which were not downloaded.
- That's it! Congrats!


## **WARNING** 
- There may be complications when running this on a non-windows machine!!
- Also, do not stop it halfway! It will leave a process running in the background.
If you are forced to stop the application, you must manually kill the process with task manager, or by restarting the computer.


### Things to add
- Let a user provide name of the excel file. (Currently has to be named GRI_2017_2020)
- Provide support for other formats than specific am and al positions in the excel.
- Error handling for wrong format of excel file.
- Allow a programmer to disable the semaphores.
- Provide better support for non-windows (The marshall.comObjectRelease only works on windows).
- Replace the current writing system with a singleton NoteTaker object.

### Other things that could be cool
- A UI system for easier targeting of the folder containing the excel file
- A UI system for targeting the excel file.
- A UI system for targeting a download folder.






#### Author
- Joachim G. Frank