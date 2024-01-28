-----------------------------------
------- Date: 27.01.2024 ----------
-----------------------------------

STALKER GAMMA DASHBOARDS (demo)

This project consists of basic data analyssis of weapons, armors and loot table in STALKER GAMMA project. Stalker GAMMA is a massive project developed by Grikitach to the game STALKER Anomaly that aims to change and overhaul the gameplay experience from base game. You can see the project portfolio here: https://github.com/Grokitach/Stalker_GAMMA. From this project you might find: what npc's player need to farm in order to get a item, how many there are in weapons and armors in games, what armor type can certain faction etc. 

The report file is stored in ~/Report/_Gamma Report.pdf or ~/Report/_Gamma Report.xlsm.

---------------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------------

The data sources are mainly two:
a) loadouts - the source code from GAMMA project;
b) weapons and armors table (refered as "discord file") - the file that is used in GAMMA discord and it stores data about all items in game

---------------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------------

The technologies used for creation of the project are:
a) VBA (Visual Basic Application) language [for binds and slicers to images],
b) Power Query (for data transformation and clearing),
c) Microsoft Excel 2021 desktop version (as a programme in which the dashboard has been made).
   It is also worth noting that this file should be older versions "friendly", there has not been used any newer functions (like unique; let; lambda; choosecols; scan; filter; xlookup; sort etc.)

!THE DASHBOARD FILE (.xlsm) IN ORDER TO WORK REQUIRES THE MACROS TO BE ENABLED! In case you do not want to turn on macros, you can also check the report pdf version (alas the slicers shall not work in it's case).

As for today, dashboard serves only DEMONSTRATION purposes - the work is in progress for more sophisticated one versions.