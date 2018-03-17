var express = require('express');
var router = express.Router();
var o = require('odata');

router.get('/contracts/:maId', function(req, res, next) {
    var id = req.params.maId;
    var oHandler = o('http://services.odata.org/V4/(S(wptr35qf3bz4kb5oatn432ul))/TripPinServiceRW/People');
    oHandler.get(function(data) {
        var clients = [];
        data.forEach(function(item){
            var i = { "UserName" : item.FirstName , "LastName" : item.LastName };
            clients.push(i);
        });
        res.send(clients);
    });
});

module.exports = router;