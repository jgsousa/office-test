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
            peopleTable.name = "PeopleTable";

            peopleTable.getHeaderRowRange().values = 
                [["First Name", "Last Name"]];
            data.forEach(function(item){
                peopleTable.rows.add(null,
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