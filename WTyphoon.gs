var ss = SpreadsheetApp.getActiveSpreadsheet();

var sheetData   = ss.getSheetByName('data');
var sheetConfig = ss.getSheetByName('config');

var numRowData = sheetData.getLastRow();

var config = {};

config.DGPA         = getConfig(2);
config.LineNotify   = getConfig(3);
config.RowIndex     = {};

config.LineNotify.Headers = { Authorization: 'Bearer ' + config.LineNotify.Token };

var countryList = getCountryList();

/* Others Functions */

function getCountryList(){

    return sheetData.getRange(2, 1, numRowData-1).getValues().map(function(item, index, array){

        config.RowIndex[item[0]] = index+2 ;
        return item[0];

    });

}

function getConfig(rowIndex){

    return JSON.parse( sheetConfig.getRange(rowIndex, 2).getValue() );

}

/* Crawlling Functions */

function xmlPreprocess(xml){

    return xml.replace(/vAlign=center/g, "vAlign='center'")
                .replace(/align=middle/g, "align='middle'")
                .replace(/align=left/g, "align='left'")
                .replace(/<FONT color=#000080 >/g, "")
                .replace(/<FONT color=#000000 >/g, "")
                .replace(/<FONT color=#FF0000 >/g, "")
                .replace(/<\/FONT>/g, "")
                .replace(/<FONT>/g, "")
                .replace(/<FONT >/g, "")
                .replace(/<br>/g, "<br/>")
                .replace(/colspan=3/g, "colspan='3'")
                .replace(/colspan=2/g, "colspan='2'")
                .replace(/:/g, "：")
                .replace(/'/g, '"')
                ;

}

function crawler() {

    var fetchURL = config.DGPA.URL;

    var responseText = UrlFetchApp.fetch(fetchURL).getContentText('UTF-8');

    // Only get table dom
    var beginSearchStr= '<TBODY class="Table_Body">';
    var endSearchStr= "</TBODY>";

    var begin = responseText.search(beginSearchStr);
    var end = responseText.search(endSearchStr);

    var xml = "<root>" + responseText.substring(begin + beginSearchStr.length, end) + "</root>";
    xml = xmlPreprocess(xml);

    var root = XmlService.parse(xml).getRootElement();

    var listCountry= root.getChildren('TR');

    var countryObject = {};

    countryList.forEach(function(country, index, array){

        countryObject[country] = {};

        var rowIndex = config.RowIndex[country];

        countryObject[country].oldContent       =  sheetData.getRange(rowIndex, 2).getValue();
        countryObject[country].newContent       =  "無颱風停班停課訊息";

    });

    if(listCountry.length > 1){

        listCountry.forEach(function(trListCountry, index, array){

            var tdListCountry = trListCountry.getChildren('TD');

            if(tdListCountry.length > 1){

                var country = tdListCountry[0].getText();

                var countryContent = tdListCountry[1].getText().replace(/。/g, "\n");

                countryObject[country].newContent = countryContent;

            }

        });

    }

    countryList.forEach(function(country, index, array){

        var rowIndex = config.RowIndex[country];

        var oldContent      = countryObject[country].oldContent;
        var newContent      = countryObject[country].newContent;

        if(oldContent !== newContent){

            sheetData.getRange(rowIndex, 3).setValue(1);

        }

        sheetData.getRange(rowIndex, 2).setValue(newContent);

    });

}

/* Notify Functions */
function checkNotify(){

    countryList.forEach(function(country, index, array){

        var rowIndex = config.RowIndex[country];

        var ifToNotify = sheetData.getRange(rowIndex, 3).getValue();

        if(ifToNotify === 1){

            toNotify(country);
            sheetData.getRange(rowIndex, 3).setValue(0);

        }

    });

}

function toNotify(country){

    var rowIndex = config.RowIndex[country];

    var message = "\n<"+country+" 停班停課狀態更新通知>\n\n"+sheetData.getRange(rowIndex, 2).getValue();

    UrlFetchApp.fetch(config.LineNotify.Endpoint, {
        headers: config.LineNotify.Headers,
            method: 'post',
            payload: {
                message : message
            }
    });

}

/* Main Rountine Function */
function main(){

    crawler();
    checkNotify();

}
