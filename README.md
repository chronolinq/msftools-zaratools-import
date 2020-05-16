# msftools-zaratools-import
A Google Apps script to import data from the MSF Toolbot into the ZaraTools Roster Organizer with one click

## Adding this script to your Roster Organizer

- Create a new sheet (not spreadsheet) in the Roster Organizer called `*Toolbot`

- In cell A1 in `*Toolbot`, enter your Toolbot Sheet ID

- Add the script to the script collection

  - Navigate to Tools > Script Editor
  
  - File > New > Script file. Name it whatever you like
  
  - Copy and paste the toolbot_import.gs contents into the new script file
  
  - Add a menu item to Triggers.gs: `topMenu.addItem("Import from MSF Toolbot", "msftoolbot_import");` on the line before `topMenu.addToUi();`
  
- Save and refresh your Roster Organizer (This will close the Script Editor)

- Access the newly added script under MSF Roster Organizer > Import from MSF Toolbot (If you do not see the MSF Roster Organizer custom menu, you may have to refresh the page again

## Credits

Discord user Hazelwood#9033 for creating the MSF Toolbot

Discord user Zarathos#6866 for creating the ZaraTools Roster Organizer
