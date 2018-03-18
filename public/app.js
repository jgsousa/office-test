'use strict';

(function () {

    //$(document).ready(function () {
    //    $('#set-color').click(retrieveData);
    //});

    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#set-color').click(retrieveData);
            $('#send-data').click(sendData);
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
            const peopleTable = currentWorksheet.tables.add("A1:C1", true /*hasHeaders*/);
            const people2Table = currentWorksheet.tables.add("A3:C3", true /*hasHeaders*/);
            peopleTable.name = "PeopleTable";
            people2Table.name = "People2Table";

            peopleTable.getHeaderRowRange().values = 
                [["First Name", "Last Name", "Value"]];
            people2Table.getHeaderRowRange().values = 
                [["First Name2", "Last Name2", "Value 2"]];

            data.forEach(function(item){
                peopleTable.rows.add(null,
                    [[ item.UserName, item.LastName, 3 ]]);
                people2Table.rows.add(null,
                    [[ item.UserName, item.LastName, 2 ]]);
            });

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function sendData() {
        Excel.run(function (context) {
            var tableRows = [];
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            const peopleTableRows = currentWorksheet.tables.getItem('people2Table').rows;
            peopleTableRows.load('items');
            return context.sync().then(function(){
                $('#listbox').addClass('show');
                var list = $('#list').empty();
                for (var i = 0; i < peopleTableRows.items.length; i++)
                {
                    var it = peopleTableRows.items[i];
                    console.log(JSON.stringify(it.values));
                    list.append(buildItem(it.values));
                    //console.log(JSON.stringify(peopleTableRows.items[i].values[0][2]));
                }
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    function buildItem(item){
        var html = '<li class="ms-ListItem" tabindex="0"><span class="ms-ListItem-primaryText">' + 
        item[0][0] + ' ' + item[0][1] +
        '<span class="ms-ListItem-secondaryText">' + 
        'Valor - ' + item[0][2] + '</span></li>';
        return html;
    }
})();