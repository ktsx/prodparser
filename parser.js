var request = require('request'),
    cheerio = require('cheerio'),
    iconv   = require('iconv-lite'),
    async   = require('async'),
    XLSX    = require('xlsx'),
    find = require('cheerio-eq');

var vendors   = [],
    config = {
    fileName  : "example.xlsx",
    output    : "out.xlsx",
    listName  : "data_export",
    baseUrl   : "http://www.abtoys.ru",
    encoding  : "win1251",
    num       : 346,
    offset    : 3,
    vendorCell: "E"
};


var workbook = XLSX.readFile(config['fileName']);

function saveVendors(config) {
    for (var i = config['offset']; i <= config['num']; i++) {
        var cell = config['vendorCell'] + i;
        vendors.push(workbook.Sheets[config['listName']][cell]['v']);
    }
    console.warn("\n=========================");
    console.error("\tVendors num - " + vendors.length);
    console.warn("=========================\n");

}
function writeDataToExcel(config, counter, data) {
    for (var i = 0; i < data.length; i++) {
        workbook.Sheets[config['listName']][data[i]['cell'] + (counter + config['offset'])]['v'] = data[i]['value'];
    }
    console.log("Filled out № - " + (counter + config['offset']));
}

function parsing(search_options, sonum, cb, coord) {
    request(search_options, function (error, response, body) {
        if (!error && response.statusCode == 200) {

            body = iconv.decode(new Buffer(body), config['encoding']);
            $    = cheerio.load(body);

            var page_link = search_options['linkstyle']();


            if (typeof(page_link) != "undefined") {
                page_link = config['baseUrl'] + page_link;

                var new_options = {
                    uri     : page_link,
                    encoding: null,
                    //proxy   : "http://122.193.14.106:82",
                    headers : {
                        'User-Agent': "ZGKMozilla " + coord.toString() + "adbc"
                    }
                };
                var blk_num = $('#container .box1');
                if(blk_num.length == 1){
                    request(new_options, function (error, response, body) {
                        if (!error && response.statusCode == 200) {
                            body = iconv.decode(new Buffer(body), config['encoding']);
                            $    = cheerio.load(body);

                            var image = $('.fot31 img').attr('src');
                            if (image == "") {
                                image = "EMPTY FIELD";
                            }else{
                                image = config.baseUrl + image;
                            }

                            var desc = $('#s1').text().trim();
                            desc     = desc.replace(/\s{2,}/g, ' ');
                            var brand = $('#s1 p').first().text();
                            brand = brand.split(": ");
                            brand = brand[1];

                            var sex = find($, '#s1 p:eq(3)').text();
                            sex = sex.split(": ");
                            sex = sex[1];
                            var country = find($, '#s1 p:eq(6)').text();
                            country = country.split(": ");
                            country = country[1];
                            var age = find($, '#s1 p:eq(2)').text();
                            age = age.split(": ");
                            age = age[1];
                            var data = [
                                {
                                    cell : "G",
                                    value: image
                                },
                                {
                                    cell : "I",
                                    value: desc
                                },
                                {
                                    cell : "K",
                                    value: brand
                                },
                                {
                                    cell : "M",
                                    value: sex
                                },
                                {
                                    cell : "L",
                                    value: country
                                },
                                {
                                    cell : "N",
                                    value: age
                                }
                            ];

                            writeDataToExcel(config, counter, data);
                            setTimeout(function () {
                                XLSX.writeFile(workbook, config['output']);
                                counter++;
                                setTimeout(function(){cb()}, 2500);
                            }, 0);
                        }
                        else {
                            console.log("[" + counter + "] - Empty");
                            counter++;
                            setTimeout(function(){cb()}, 2500);
                        }

                    });
                } else{
                    console.log("[" + counter + "] - Not found best variant");
                    counter++;
                    setTimeout(function(){cb()}, 2500);
                }
            } else {
                //TODO 1)check num of search options 2)if ok then call parsing width new options, else callback()
                console.warn("Undefined url: [" + (counter + config['offset']) + "]");
                if (sonum < search_options.length) {
                    sonum++;
                    parsing(search_options[sonum]);
                } else {
                    counter++;
                    setTimeout(function(){cb()}, 2500);
                }
            }


        } else {
            console.log("[" + counter + "] - Empty");
            counter++;
            setTimeout(function(){cb()}, 2500);
        }

    });
}

//Collect vendors
saveVendors(config);

var counter = 0;

async.whilst(function exitCondition() {
        return counter < vendors.length;
    },
    function increaseCounter(cb) {
        var sonum          = 0;
        var coord          = Math.floor(Math.random() * (75 - 5 + 1)) + 5;
        var search_options = [
            {
                uri      : 'http://www.abtoys.ru/Go/Products?keywords=' + encodeURIComponent(vendors[counter])+"&x=13&y=21",
                encoding : null,
                //proxy   : "http://122.193.14.106:82",
                headers  : {
                    'User-Agent': "HMMozilla " + coord.toString() + "ab"
                },
                linkstyle: function () {
                    return $('.img_toy a').attr('href');
                }
            }
        ];
        parsing(search_options[sonum], sonum, cb, coord);

    },
    function cb(err) {
        if (err) {
            console.log(err);
            return;
        }
        console.log("Job complete");

    }
);