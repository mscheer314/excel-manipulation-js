let Excel = require("exceljs");
const workbook = new Excel.Workbook();
const rolesFromScript = [];
const rolesToCompare = [];

workbook.xlsx.readFile("newSheet.xlsx").then(function () {
    const powerChartWorksheet = workbook.getWorksheet("PowerChartExport");
    const umraWorksheet = workbook.getWorksheet("UMRA");

    // get roles created from recreateRoleListingFromUMRA.js
    umraWorksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        rolesFromScript.push({
            lastName: row.findCell(1).text,
            firstName: row.findCell(2).text,
            role: row.findCell(8).value,
        });
    });

    let countWithPowerChartRole = 0;

    rolesFromScript.forEach((element) => {
        powerChartWorksheet.eachRow(
            { includeEmpty: false },
            function (row, rowNumber) {
                if (
                    // element.firstName is 'firstName' everytime?!?!?! WHY?????
                    row.findCell(1).text.includes(element.firstName) &&
                    row.findCell(1).text.includes(element.lastName)
                ) {
                    element.powerChartRole = row.findCell(2).text;
                }
            }
        );
    });

    // get all of the roles from UMRA into rolesToCompare
    rolesFromScript.forEach((element) => {
        if (
            // if rolesToCompare does not already contain the role entry
            !rolesToCompare.some((item) => item.role === element.role)
        ) {
            rolesToCompare.push({
                role: element.role,
                // create empty array to put any PowerChart roles into later
                powerChartRoles: [],
            });
        }
    });

    // look for PowerChart roles to put into the results array
    rolesFromScript.forEach((element) => {
        rolesToCompare.forEach((item) => {
            if (element.firstName === "First Name") {
                return;
            }

            if (
                element.role === item.role &&
                !item.powerChartRoles.includes(element.powerChartRole) &&
                element.powerChartRole !== undefined
            ) {
                item.powerChartRoles.push(element.powerChartRole);
                countWithPowerChartRole++;
            }
        });
    });
    //console.log(rolesToCompare);
    const resultsWorksheet = workbook.addWorksheet("results");

    for (let i = 0; i < rolesToCompare.length; i++) {
        resultsWorksheet.addRow([
            rolesToCompare[i].role,
            rolesToCompare[i].powerChartRoles,
        ]);
    }

    return workbook.xlsx.writeFile("finalSheet.xlsx");
});
