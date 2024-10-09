# BLIND GANTT

## title: SCOPE

This application aims to support visually impaired people to be able to generate a simple gantt graph by filling an excel file with a predefined format.

## Features

The main features in this initial version are:
    1. Specify tasks in an excel file
    2. Each task is defined by its ID, name, start date, duration in hours and predecessors
    3. Task ID and predecessor must be a number
    4. Start date format is dd/mm/yyyy
    5. The weekd-end is assumed to be on saturday and sunday
    6. Official holidays are got automatically based on the country location obtained from system  
    
## Application outputs

After execution, the application provide the following outputs:     
    1. Update the Excel file to align when necessary start dates according to predecessors end dates
    2. Save  the Gantt diagram on the hard disk with a ".png" file extension

``Note: Obtaining  national holidays list could be not successful, in that case, the Gantt is geerated without taking them in consideration ``

## Files sctructure

In the project folder you have several files. This is below their signification:
1. blind_gantt.py : ths is the console application file
2. blind_gantt_visual.py : this is the visual application that should generate a graphical interface and allow to see a real-time update of the Gantt graph when we modify the task lisk in the data grid. The application allow also to save the data grid as an excel sheet and the Gantt graph as a png picture.
3. lod.txt : this file will log each error occuring during executionsuch the unability to obtain national holidays list. If no error occurs during execution, a  line specifying that the execution is successful will be wrote in it.
4. Dist folder : this folder contain the executable to deploy.
5. Other files and  folders are related to the environment and libraries 

``Note: The Excel file shall be in the same folders with the executable to avoid a wrong behavior. ``