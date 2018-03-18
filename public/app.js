'use strict';

(function () {

    //$(document).ready(function () {
    //    $('#set-color').click(retrieveData);
    //});

    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#set-color').click(retrieveData);
        });
    };

    function retrieveData() {
        var contract = $('.ms-TextField-field').val();
        $.ajax({
            url: "/api/contracts/" + contract,
            cache: false,
            success: function(data){
                setExcelData(data);
            }
          });
    }

    function setExcelData(data) {
        Excel.run(function (context) {
            console.log(JSON.stringify(data));
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            const peopleTable = currentWorksheet.tables.add("A1:B1", true /*hasHeaders*/);
            const people2Table = currentWorksheet.tables.add("A3:B3", true /*hasHeaders*/);
            peopleTable.name = "PeopleTable";
            people2Table.name = "People2Table";

            peopleTable.getHeaderRowRange().values = 
                [["First Name", "Last Name"]];
            people2Table.getHeaderRowRange().values = 
                [["First Name2", "Last Name2"]];

            data.forEach(function(item){
                peopleTable.rows.add(null,
                    [[ item.UserName, item.LastName ]]);
                people2Table.rows.add(null,
                    [[ item.UserName, item.LastName ]]);
            });

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
})();