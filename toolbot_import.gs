var toolbot_id = 0;
var toolbot_power = 7;
var toolbot_star = 8;
var toolbot_redstar_active = 9;
var toolbot_redstar_inactive = 23;
var toolbot_level = 10;
var toolbot_gear = 11;
var toolbot_basic = 18;
var toolbot_special = 19;
var toolbot_ultimate = 20;
var toolbot_passive = 21;
var toolbot_shards = 22;

var toolbot_gear_topleft = 12;
var toolbot_gear_middleleft = 13;
var toolbot_gear_bottomleft = 14;
var toolbot_gear_topright = 15;
var toolbot_gear_middleright = 16;
var toolbot_gear_bottomright = 17;

function msftoolbot_import()
{
    // Get Toolbot Sheet Id
    var toolbotSheetId = SpreadsheetApp.getActive().getSheetByName('*Toolbot').getRange("A1").getValue();

    // Get Toolbot Range and Values
    var rosterRange = SpreadsheetApp.openById(toolbotSheetId).getSheetByName("Roster").getRange("A2:X200");
    var roster = rosterRange.getValues();

    // Hydrate the id <> name dictionary
    getNameDictionary();
    // Get the names and data from the Roster sheet
    var names = getNamedRangeA("Roster_Names").getValues();
    var data = getNamedRangeA("Roster_Data").getValues();
    var farmingNames = getNamedRangeA("Farming_Names").getValues();
    var farmingGear = getNamedRangeA("Farming_GearParts").getValues();
    for (var i = 0; i < roster.length; i++)
    {
        // Try and get the Name from the id on on the row
        var currentRow = roster[i];
        var id = currentRow[toolbot_id];
        var name = nameOf(id);

        // If there's no name found, skip
        if (name)
        {
            // Find the row
            var s = -1;
            for (var r = 0; r < names.length; r++)
            {
                if (names[r][0] == name)
                {
                    s = r;
                    break;
                }
            }

            // If this character is on the roster sheet, write out their info
            if (s > -1)
            {
                var currentData = data[s];

                currentData[0] = currentRow[toolbot_star];
                currentData[1] = currentRow[toolbot_redstar_active] + currentRow[toolbot_redstar_inactive];
                currentData[2] = currentRow[toolbot_level];
                currentData[3] = currentRow[toolbot_basic];
                currentData[4] = currentRow[toolbot_special];
                currentData[5] = currentRow[toolbot_ultimate];
                currentData[6] = currentRow[toolbot_passive];
                currentData[7] = currentRow[toolbot_gear];
                currentData[8] = currentRow[toolbot_power];
            }

            // Try and fill in the farming tab
            var farmingRow = -1

            for (var f = 0; f < farmingNames.length; f++)
            {
                if (farmingNames[f][0] == name)
                {
                    farmingRow = f;
                    break;
                }
            }

            // If this character is in the farming tab, write out their equipped gear
            if (farmingRow > -1)
            {
                var currentFarming = farmingGear[farmingRow];

                currentFarming[0] = (currentRow[toolbot_gear_topleft] == "Y") ? "TRUE" : "FALSE";
                currentFarming[1] = (currentRow[toolbot_gear_middleleft] == "Y") ? "TRUE" : "FALSE";
                currentFarming[2] = (currentRow[toolbot_gear_bottomleft] == "Y") ? "TRUE" : "FALSE";
                currentFarming[3] = (currentRow[toolbot_gear_topright] == "Y") ? "TRUE" : "FALSE";
                currentFarming[4] = (currentRow[toolbot_gear_middleright] == "Y") ? "TRUE" : "FALSE";
                currentFarming[5] = (currentRow[toolbot_gear_bottomright] == "Y") ? "TRUE" : "FALSE";
            }
        }
    }

    getNamedRangeA("Roster_Data").setValues(data);
    getNamedRangeA("Farming_GearParts").setValues(farmingGear);
}