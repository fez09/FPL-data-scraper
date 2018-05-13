# Background
I am not a programmer. I do this as a way to learn more about programming. Being an amateur the code for this script is very raw and has plenty of things which can definitely be improved, and the code is not 'clean'. That said, I am satisfied with how this turned out. Will gladly accept advice/pointers from more experienced people to improve this code. Lets get to it. 

I enjoy the barclays premier league and play the fantasy a lot. One thing I did not like was the fact that you are not able to view detailed history of past seasons. So I started logging everything manually in excel sheets. Decided it would be easier to write a program for it. I chose python since I just started learning it a couple of months ago and the openpyxl module is a way to export data to excel sheets. I'm aware there is a possibility something like this already exists considering the number of 3rd party statistics sites available. But I wanted to make my own script. 
# The Script
Executing the script prompts a GUI interface asking the user to enter their fantasy team ID. This can be found easily on the wesite. Clicking the submit button then proceeds to import all json data from the website and exporting it to excel. 

Modules used are mainly tkinter for GUI, openpyxl for Excel shenanigans and requests to fetch Json data. 

# List of data imported
 - Gameweek Score
 - Gameweek Average Score
 - Points Benched
 - Transfers Made for GW
 - Transfer Cost
 - Gameweek Rank
 - Overall Points
 - Overall Rank
 - Position in overall leaderboard
 - Team Value
 - Weekly Squad
 - Chips Used and when
 - Complete Transfer history with values
 - Final Dream Team
 - Cup history - Added on 13 May 2018
# Json links
 - https://fantasy.premierleague.com/drf/entry/{}/event/{}/picks (FPL ID, Gameweek number) - Live team points
 - https://fantasy.premierleague.com/drf/entry/{}/history (GW number) - GW history
 - https://fantasy.premierleague.com/drf/bootstrap-static  - Contains all player data
 - https://fantasy.premierleague.com/drf/entry/{}/transfers (GW number) - Transfer history
 - https://fantasy.premierleague.com/drf/dream-team  - Dream team
 - https://fantasy.premierleague.com/drf/entry/{}/cup - Cup History - Added on 13 May 2018
