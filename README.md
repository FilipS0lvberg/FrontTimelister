# FrontTimelister
Repository for Front Generation - Timelister
The code is written quick and dirty (to save costs), but works as intended.
Should probably be refactored and deployed to a production server if frequently used and bugs are discovered - streamlit should be sufficient for this application.
There are some performance issues related to the download buttons and rendering of multiple of these buttons simultaneously (which happens if you write the password (authenticate) in the first page.)

The structure of code is the following:
1. Main file: FrontTimelister.py (this file runs all the code and imports the necessary functions from the other code files)
2. supporting_functions.py - is a small library of functions that execute small specific tasks. The other files import some of these specific functions. 
3. create_prosjektregnskap.py - this code contains a single function: create_prosjektregnskap1 which is responsible for generating a prosjekt regnskap excel file 
and converting it to a bytes IO format so that streamlit can upload it to the browser. The code imports the necessary functions from the supporting functions file.
4. create_timelsiter.py - this code contains a single function: create_timelister1 which is responsible for generating a timeliste excel file and converting it to a bytes IO format so that stramlit can upload it to the browser. The code imports the necessary functions from the supporting functions file.
5. write_erklering.py - this code contains a single function: write_erklering1 which is responsible for generating a erklering word file and converting it to a bytes IO format so that stramlit can upload it to the browser.
6. full_download.py - this code contains a single function responsible for taking in multiple bytes IO objects and adding it to a compressed folder (zip) which also is in bytes IO format so that streamlit can upload it to the browser.
7. requirements.txt - this file contains the necessary python packages for running the program in streamlit community cloud. Note that this is not sufficient to run on a different system/ machine. A minimum requirement is to have streamlit installed in the environment also - there may be other package requirements as well. Python 3.10 is used for this prosject.

Command for running the code in terminal (or other): streamlit run FrontTimelister.py
