// -------------------------------------------------
// MSF Tools Import
// Author: chronolinq
// -------------------------------------------------

// Range Values
let tb_sheet_name = "*Toolbot";
let tb_id_cell = "A1";
let tb_roster_sheet_name = "Roster";
let tb_roster_sheet_range = "A2:X";
let r_name_range = "Roster_Names";
let r_data_range = "Roster_Data";
let r_farming_name_range = "Farming_Names";
let r_farming_gear_range = "Farming_GearParts";

// Roster Organizer Data Columns
let r_star = 0;
let r_redstar = 1;
let r_level = 2;
let r_basic = 3;
let r_special = 4;
let r_ultimate = 5;
let r_passive = 6;
let r_gear = 7;
let r_power = 8;
let r_shards = 9;

// Roster Organizer Gear Columns
let r_gear_topleft = 0;
let r_gear_middleleft = 1;
let r_gear_bottomleft = 2;
let r_gear_topright = 3;
let r_gear_middleright = 4;
let r_gear_bottomright = 5;

// Toolbot Data Columns
let tb_id = 0;
let tb_power = 7;
let tb_star = 8;
let tb_redstar_active = 9;
let tb_redstar_inactive = 23;
let tb_level = 10;
let tb_gear = 11;
let tb_basic = 18;
let tb_special = 19;
let tb_ultimate = 20;
let tb_passive = 21;
let tb_shards = 22;

// Toolbot Gear Columns
let tb_gear_topleft = 12;
let tb_gear_middleleft = 13;
let tb_gear_bottomleft = 14;
let tb_gear_topright = 15;
let tb_gear_middleright = 16;
let tb_gear_bottomright = 17;

// Keywords
let kw_clear_value = "None";
let kw_max_shards = "MAX";
let kw_equipped = "Y";

function msftoolbot_import()
{
    // Get Toolbot Sheet Id
    let toolbotSheetId = SpreadsheetApp.getActive().getSheetByName(tb_sheet_name).getRange(tb_id_cell).getValue();

    // Get Toolbot Range and Values
    let tbRange = SpreadsheetApp.openById(toolbotSheetId).getSheetByName(tb_roster_sheet_name).getRange(tb_roster_sheet_range);
    let tbData = tbRange.getValues();

    // Hydrate the id <> name dictionary
    getNameDictionary();

    // Get the names and data from the Roster sheet
    let rNames = getNamedRangeA(r_name_range).getValues();
    let rData = getNamedRangeA(r_data_range).getValues();
    let farmingNames = getNamedRangeA(r_farming_name_range).getValues();
    let farmingGear = getNamedRangeA(r_farming_gear_range).getValues();

    // Look through each row in the Toolbot range
    for (let i = 0; i < tbData.length; i++)
    {
        // Try and get the Name from the Id on the row
        let currentTbRow = tbData[i];
        let id = currentTbRow[tb_id];
        let name = nameOf(id);

        // If there's no name found in dictionary, skip
        if (name)
        {
            // If this character is on the roster sheet, write out their info
            let rRow = getRowByName(rNames, name);

            if (rRow > -1)
            {
                let currentData = rData[rRow];

                setRosterValue(currentData, currentTbRow, r_level, tb_level);
                setRosterValue(currentData, currentTbRow, r_star, tb_star);
                setRosterRedstarValue(currentData, currentTbRow, r_redstar, tb_redstar_active, tb_redstar_inactive);
                setRosterValue(currentData, currentTbRow, r_power, tb_power);
                setRosterValue(currentData, currentTbRow, r_basic, tb_basic);
                setRosterValue(currentData, currentTbRow, r_special, tb_special);
                setRosterValue(currentData, currentTbRow, r_ultimate, tb_ultimate);
                setRosterValue(currentData, currentTbRow, r_passive, tb_passive);
                setRosterValue(currentData, currentTbRow, r_gear, tb_gear);
                setShardValue(currentData, currentTbRow, r_shards, tb_shards);
            }

            // If this character is in the farming tab, write out their equipped gear
            let fRow = getRowByName(farmingNames, name);

            if (fRow > -1)
            {
                let currentFarming = farmingGear[fRow];

                setFarmingValue(currentFarming, currentTbRow, r_gear_topleft, tb_gear_topleft);
                setFarmingValue(currentFarming, currentTbRow, r_gear_middleleft, tb_gear_middleleft);
                setFarmingValue(currentFarming, currentTbRow, r_gear_bottomleft, tb_gear_bottomleft);
                setFarmingValue(currentFarming, currentTbRow, r_gear_topright, tb_gear_topright);
                setFarmingValue(currentFarming, currentTbRow, r_gear_middleright, tb_gear_middleright);
                setFarmingValue(currentFarming, currentTbRow, r_gear_bottomright, tb_gear_bottomright);
            }
        }
    }

    // Set the updated data back on the ranges
    getNamedRangeA(r_data_range).setValues(rData);
    getNamedRangeA(r_farming_gear_range).setValues(farmingGear);

}

// Set the Toolbot values on the Roster Organizer
function setRosterValue(rData, tbData, rIndex, tbIndex)
{
    let value = tbData[tbIndex];

    if (value == kw_clear_value) rData[rIndex] = "";
    else if (value) rData[rIndex] = value;
}

// Set the Toolbot redstar values on the Roster Organizer
function setRosterRedstarValue(rData, tbData, rIndex, tbRedstarActiveIndex, tbRedstarInactiveIndex)
{
    // The MSF Toolbot records active and inactive redstars in separate columns so they must be added
    let redstar = tbData[tbRedstarActiveIndex]
    let inactive = tbData[tbRedstarInactiveIndex]

    if (redstar == kw_clear_value) 
    {
        // Clear the value if Red Stars is "None"
        rData[rIndex] = "";

        // Set the red star value to the inactive value if it's not null and it's not "None"
        if (inactive && inactive != kw_clear_value)
        {
            rData[rIndex] = inactive;
        }

        return;
    }
    else if (redstar || inactive)
    {
        // If we have a value for inactive, and it's not "None", add it to Red Star
        if (inactive && inactive != kw_clear_value)
        {
            redstar += inactive;
        }

        rData[rIndex] = redstar;
    }
}

// Set the Toolbot shard values on the Roster Organizer
function setShardValue(rData, tbData, rIndex, toolbotIndex)
{
    // If the character's shards are maxed out, the value will be "MAX". The value should be cleared in that case
    let value = tbData[toolbotIndex];

    if (value == kw_clear_value || value == kw_max_shards) rData[rIndex] = "";
    else if (value) rData[rIndex] = value;
}

// Set the Toolbot equipped gear values on the Roster Organizer
function setFarmingValue(rData, tbData, rIndex, tbIndex)
{
    // If equipped, the value will be "Y", otherwise it will be "N"
    rData[rIndex] = (tbData[tbIndex] == kw_equipped) ? "TRUE" : "FALSE";
}

// Given a name, find the row that, that name exists
function getRowByName(names, name)
{
    let row = -1;

    for (let i = 0; i < names.length; i++)
    {
        if (names[i][0] == name)
        {
            row = i;
            break;
        }
    }

    return row;
}