# Fantasy Premier League Data Fetcher
Code was written in Python 3. I chose python since I just started learning it a couple of months ago and the openpyxl module is a way to export data to excel sheets.

Executing the script prompts a GUI interface asking the user to enter their fantasy team ID. This can be found easily on the wesite. Clicking the submit button then proceeds to import all json data from the website and exporting it to excel. 

Youtube tutorial: https://youtu.be/DM9-e84iDZ8

Screenshots: https://imgur.com/a/lp27xAn

## **Data updated for 2020/2021 Season**

## Requirements
Whatever is required by the openpyxl and tkinter modules, ie,
 - Windows device to run the executable
 
## Modules Used 
 - tkinter - for GUI
 - requests - to fetch json data from FPL website
 - openpyxl - to write the imported data to excel workbook
 - os - to show the path of the created excel file
 - I've also used urlopen and base64 to read the FPL logo off of an imgur link to place it in the GUI 
 
## List of data imported
 - Gameweek Score
 - Gameweek Average Score
 - Points Benched
 - Transfers Made for GW
 - Transfer Cost
 - Gameweek Rank
 - Overall Points
 - Overall Rank
 - Percentile Overall Rank (v2.0+)
 - Percentile Gameweek Rank (v2.0+)
 - Position in overall leaderboard
 - Team Value
 - Weekly Squad
 - Chips Used and when
 - Complete Transfer history with values
 - Final Dream Team
 - FPL Cup history
 - Classic League Ranks 
 - H2H Ranks
 - Individual player scores for each GW
 - Weekly Dream Team
 
## Json links
 - http://fantasy.premierleague.com/drf/entry/{}/event/{}/picks (FPL ID, Gameweek number) - Live team points
 - http://fantasy.premierleague.com/drf/entry/{}/history (FPL ID) - GW history
 - http://fantasy.premierleague.com/drf/bootstrap-static  - Contains all player data
 - http://fantasy.premierleague.com/drf/entry/{}/transfers (GW number) - Transfer history
 - http://fantasy.premierleague.com/drf/dream-team  - Dream team
 - http://fantasy.premierleague.com/drf/entry/{}/cup - Cup History
 - https://fantasy.premierleague.com/drf/event/{}/live (FPL ID) - Live Player Points
 - https://fantasy.premierleague.com/drf/dream-team/{} (GW Number) - Weekly Dream Team

_____________________________________________________________________________________________________________________
www.fezfiles.net
