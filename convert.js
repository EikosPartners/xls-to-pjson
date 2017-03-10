const xlsx = require('xlsx');
const fsw = require('ep-utils/fsWrapper');

/**
 * Method to read in an excel spreadsheet.
 *
 * @param filename : String - The name of the file to read
 * @param output : String - The name of the json output file
 * @param headerRow : Number - The row corresponding to the header of the sheet, defaults to 0
 */
function read(filename, output, headerRow = 0) {
    let file = xlsx.readFile(filename),
        actions = [];

    let sheetNames = file.SheetNames;

    console.log(sheetNames);

    sheetNames.forEach( (name) => {
        let sheet = file.Sheets[name];

        let numRows = sheet["!range"].e.r;

        for (let i = headerRow + 1; i <= numRows; i++) {
            let action = {
                type: "action",
                actionType: "ajax",
                options: {
                    target: {
                        uri: "",
                        name: "",
                        options: {
                            headers: {

                            }
                        }
                    },
                    params: {

                    }
                }
            };

            let A = "A" + i,
                B = "B" + i,
                C = "C" + i,
                D = "D" + i;

            if (sheet[A]) {
                action.options.target.name = sheet[A].v;
            }

            if (sheet[B]) {
                action.options.target.uri = sheet[B].v;
            }

            try {
                if (sheet[C]) {
                    action.options.params = JSON.parse(sheet[C].v);
                }
            } catch (e) {
                console.error(e);
            }

            try {
                if (sheet[D]) {
                    action.options.target.options.headers = JSON.parse(sheet[D].v);
                }
            } catch (e) {
                console.error(e);
            }

            actions.push(action);
        }

        write(output, actions);
    });
}

/**
 * Method to write the json to a file.
 */
function write(filename, data) {
    fsw.jsonToFile(filename, data, (err, res) => {
        if (err) {
            console.log(err);
        } else {
            console.log(res);
        }
    });
}

module.exports = {
    read: read,
    write: write
};
