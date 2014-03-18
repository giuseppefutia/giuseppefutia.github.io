var office = require('office');
var fs = require('fs');

office.parse('./tasso-disoccupazione-regione.xls', function(err, data) {
    var json = data.sheets.sheet.rows.row;
    console.log(json);

    var outputFilename = './tasso-disoccupazione-regione.json';
    fs.writeFile(outputFilename, JSON.stringify(json, null, 4), function(err) {
        if(err) {
            console.log(err);
        } else {
     	   console.log("JSON saved");
   		}
	});
});