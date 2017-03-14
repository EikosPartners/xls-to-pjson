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
 * @param isSeries : Boolean - Determine whether to add the actions into a series
 * @param headerRow : Number - The row corresponding to the header of the sheet, defaults to 0
 */
function flat(filename, output, isSeries, headerRow = 0) {
    let actions = buildActions(filename, headerRow, 'flat');

    if (isSeries) {
        actions = series(actions);
    }

    write(output, actions);
}

/**
 * Method to create a psjon workflow from an excel spreadsheet.
 *
 * @param filename : String - The name of the excel file
 * @param output : String - The name of the json output file
 * @param isSeries : Boolean - Determine whether to add the actions into a series
 * @param headerRow : Number - The number pertaining to the header row of the input file
 */
function workflow(filename, output, isSeries, headerRow = 0) {
    let actions = buildActions(filename, headerRow, 'workflow'),
        baseObj = actions.splice(0,1)[0];

    addNestedActions(baseObj, actions);

    if (isSeries) {
        baseObj = series([baseObj]);
    }

    write(output, baseObj);
}

/**
 * Method to convert an excel spreadsheet into ajax actions within a series
 *
 * @param actions : Array of actions to add to the series object
 *
 * @return Returns a series action with the inputted actions
 */
function series(actions) {
    let seriesJson = {
            type: "action",
            actionType: "series",
            text: "Start API Actions",
            options: {
                actions: actions
            }
        };

    return seriesJson;
}

/**
 * Method to build an array of actions from an excel spreadsheet.
 *
 * @param filename : String - The name of the excel spreadsheet.
 * @param headerRow : Number - The row corresponding to the header of the
 * @param type : String - "workflow" or "flat" to determine mergeid;
 *
 * @return Returns an array of actions
 */

function buildActions(filename, headerRow, type) {
    let file = xlsx.readFile(filename),
        actions = [],
        sheetNames = file.SheetNames;

    sheetNames.forEach( (name) => {
        let sheet = file.Sheets[name],
            numRows = sheet["!range"].e.r,
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

                action.options.mergeid = (type === 'workflow') ? "update_grid_error_actions" : "update_grid_actions";

                let A = "A" + i,
                    B = "B" + i,
                    C = "C" + i,
                    D = "D" + i,
                    E = "E" + i;

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

                if (sheet[E]) {
                    // Column E will be the method type, empty is GET.
                    action.options.target.options.type = sheet[E].v;
                }

                actions.push(action);
            }
    });

    return actions;
}

/**
 * Method to add actions to an object recursively to create a nested object
 *
 * @param obj : Object - The object to add the actions too
 * @param actions: Array - The array of actions to add
 */
 function addNestedActions(obj, actions) {
     obj.options = obj.options || {};
     // Add the keyMap.
     obj.options.keyMap = {
         resultsKey: "data"
     };
     obj.options.nextActions = [];
     obj.options.errorActions = [eventAction];
     // Add the event Action.
     obj.options.nextActions.push(eventAction);

     if (actions.length === 0) {
         return;
     }

     // Add the ajax action.
     obj.options.nextActions.push(actions.splice(0,1)[0]);

     addNestedActions(obj.options.nextActions[1], actions);
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
