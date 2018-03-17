'use strict';

(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#set-color').click(retrieveData);
        });
    };

    function retrieveData() {
        var oHandler = o('http://services.odata.org/V4/(S(wptr35qf3bz4kb5oatn432ul))/TripPinServiceRW/People');
        oHandler.get(function(data) {
            console.log(data); // data of the TripPinService/People endpoint
        });
    }

    function setExcelData() {
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