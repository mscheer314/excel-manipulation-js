let Excel = require("exceljs");
const workbook = new Excel.Workbook();
let worksheet;

// read UMRA export and grab the fields to recreate the role listing
// as it is listed in Role Management
workbook.xlsx.readFile("RoleAudit.xlsx").then(function () {
    worksheet = workbook.getWorksheet("UMRA");
    // console.log(worksheet);
    let roles = ["rolesFromScript"];

    worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (
            row.findCell(5).value === "BJ Extended Care" ||
            row.findCell(5).value === "Eunice Smith Home" ||
            row.findCell(5).value === "Village North - Skilled"
        ) {
            roles.push(
                `Role_${row.getCell(5).text.replace(/ /g, "_")}_${row
                    .getCell(6)
                    .text.replace(/ /g, "_")}_${row
                    .getCell(3)
                    .text.replace(/ /g, "_")}`
            );
        } else {
            roles.push(
                `Role_${row.getCell(4).text.replace(/ /g, "_")}_${row
                    .getCell(5)
                    .text.replace(/ /g, "_")}_${row
                    .getCell(3)
                    .text.replace(/ /g, "_")}`
            );
        }
    });

    let roleColumn = worksheet.getColumn("H");

    roleColumn.eachCell((cell, i) => {
        cell.value = roles[i];
    });

    return workbook.xlsx.writeFile("newSheet.xlsx");
});
