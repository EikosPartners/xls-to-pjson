const xlsx = require('xlsx');
const fsw = require('ep-utils/fsWrapper');

const eventAction = {
    type: "action",
    actionType: "event",
    options: {
        target: "apis_grid.add",
        params: [
            {
                "name": "{{request.name}}",
                "endpoint": "{{request.uri}}",
                "status": "{{status}}"
            }
        ],
        useOptions: true
    }
};

/**
 * Method to convert an excel spreadsheet to a flat json array.
 *
 * @param filename : String - The name of the file to read
 * @param output : String - The name of the json output file
 * @param headerRow : Number - The row corresponding to the header of the sheet, defaults to 0
 */
function flat(filename, output, headerRow = 0) {
    let file = xlsx.readFile(filename),
        actions = [];

    let sheetNames = file.SheetNames;

    console.log(sheetNames);

    sheetNames.forEach( (name) => {
        let sheet = file.Sheets[name];

        console.log(sheet);

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
 * Method to create a psjon workflow from an excel spreadsheet.
 *
 * @param filename : String - The name of the excel file
 * @param output : String - The name of the json output file
 * @param headerRow : Number - The number pertaining to the header row of the input file
 */
function workflow(filename, output, headerRow = 0) {
    let file = xlsx.readFile(filename),
        actions = [];

    let sheetNames = file.SheetNames;

    sheetNames.forEach( (name) => {
        let sheet = file.Sheets[name];
        let numRows = sheet["!range"].e.r,
            action = {
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

                    },
                    nextActions: []
                }
            };

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

            if (sheet[B]){
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

        let baseObj = actions.splice(0,1)[0];

        addActions(baseObj, actions);

        write(output, baseObj);
    });
}

/**
 * Method to add actions to an object recursively to create a nested object
 *
 * @param obj : Object - The object to add the actions too
 * @param actions: Array - The array of actions to add
 */
 function addActions(obj, actions) {
     if (actions.length === 0) {
         return;
     }

     obj.options = obj.options || {};
     obj.options.nextActions = [];
     // Add the event Action.
     obj.options.nextActions.push(eventAction);
     // Add the ajax action.
     obj.options.nextActions.push(actions.splice(0,1)[0]);

     addActions(obj.options.nextActions[1], actions);
 }

/**
 * Method to write the json to a file.
 */
function write(filename, data) {
    fsw.jsonToFile(filename, data, (err, res) => {
        if (err) {
            console.log(err);
        }
    });
}

module.exports = {
    flat: flat,
    write: write,
    workflow: workflow
};
