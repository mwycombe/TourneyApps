# TourneyApps
Home for apps used to run tourneys
## Inventory 

### TTourney_2026_Actual_Post_Mortem
This was the starting point for the 2026 version used for the April 2026 Tourney.
YB Prelim Reg has the enhance prelimReg Verification version
It empty.

### Tourney_2026_Tourney_Clean_April_Actual
This is the preserved version of the Tourney App with the values used during the tourney.
YB Prelim Reg has the enhanced prelimReg Verification version.\
This was added to do retroactive verification of prelimReg provided for actual tourney.\
**DO NOT ALTER THIS VERSION. MAKE COPIES AS NEEDED**

### Tourney_2026_Actual_Reset
This is a copy of the April_Actual used for any changes going forward.

## How to change a reference to TourneyGlobalObjects.xlsm 
This is a messy process that does not work the way you think in VBA->Tools->References.
1.  Close all files.
2.  Open the TourneyGlobalObjects.xlsm in the directory of interest.
3.  F4 for properties and make sure the name is TourneyGlobals.
4.  Now, open the target file where you want to make the reference link.
5.  Open VBA Tools/Reference and click on the TourneyGlobals name.
6.  Go to File/Options/Trust Center/Macros and make sure macros are enabled.
7.  Finally, got to Trusted Locations and select the directory where the referenced file is.
8.  And lastly, modify this trusted location to allow subfolders to be include.
9.  Save and close the target file and referenced file and reopen the target file.
10. The reference should now be resolved giving access to all the macros inside TourneyGlobals.

## Change Log
- [X] Rename Tourney_2026_Actual_Reset to Tourney_2026_Actual_Post_Mortem <br>
      Use this as a base for 2027 development\
      - [X] Copy over enhanced prelimReg from Tourney_2026_Testing_Reset_Clean
- [X] Make sure every copy uses the TourneyGlobalObjects.xlsm in the local TourneyGlobals directory.<br>   Past versions referred out to a common tourney globals outside the app directory making it non-portable.\
prelimReg does not use TourneyGlobalObjects.xlsm so nothing to change.
  - [X] Tourney_2026_Testing_Reset_Clean\
  - [X] Tourney_2026_Actual_Post_Mortem\

- [X] Copy enhanced verification code into Registration and Main Roster<br>
      MainRoster has Finalize for No Shows; ConsyRoster has Finalize Consy Entrants<br>
      Leave actual untouched as System of Record for April 2026 tourney.<br>

    - [X] Tourney_2026_Testing_Reset_Clean
    - [X] Tourney_2026_Actual_Post_Mortem

- [X] In Enhanced prelimReg remove the "ShowHide" TextFrame reference<br>    

    - [X] Tourney_2026_Testing_Reset_Clean
    - [X] Tourney_2026_Actual_Post_Mortem

- MainRoster Finalize For No Shows button just makes sure that there are no pool entries against unentered players as this would throw off the pool calculations. This is a final check as the pool cells should have been cleaned up in Registration.<br>
- ConsyRoster Finalize Consy Entrants is the same as MainRoster Finalize button. No need to undergo extensive pool cell validation as only the Equal Pool is ever used.


