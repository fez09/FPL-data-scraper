# Fantasy Premier League Data Fetcher

Executing the script prompts a GUI interface asking the user to enter their fantasy team ID. This can be found easily on the wesite. Clicking the submit button then proceeds to import all json data from the website and exporting it to excel. 

Modules used are mainly tkinter for GUI, openpyxl for Excel shenanigans and requests to fetch Json data. 

## List of data imported
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
 - FPL Cup history
 - Classic League Ranks 
 - H2H Ranks
 - Individual player scores for each GW
 - Weekly Dream Team
 
## Json links
 - http://fantasy.premierleague.com/drf/entry/{}/event/{}/picks (FPL ID, Gameweek number) - Live team points
 - http://fantasy.premierleague.com/drf/entry/{}/history (GW number) - GW history
 - http://fantasy.premierleague.com/drf/bootstrap-static  - Contains all player data
 - http://fantasy.premierleague.com/drf/entry/{}/transfers (GW number) - Transfer history
 - http://fantasy.premierleague.com/drf/dream-team  - Dream team
 - http://fantasy.premierleague.com/drf/entry/{}/cup - Cup History
 - https://fantasy.premierleague.com/drf/event/{}/live (FPL ID) - Live Player Points
 - https://fantasy.premierleague.com/drf/dream-team/{} (GW Number) - Weekly Dream Team
