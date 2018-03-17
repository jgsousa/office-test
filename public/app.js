'use strict';

(function () {

    $(document).ready(function () {
        $('#set-color').click(retrieveData);
    });

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
                console.log(JSON.stringify(data));
              //setExcelData(data);
            }
          });
    }

    function setExcelData(data) {
        Excel.run(function (context) {
            var range = context.workbook.getSelectedRange();
            range.format.fill.color = 'green';

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
})();