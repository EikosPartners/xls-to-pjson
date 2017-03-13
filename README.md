# xls-to-pjson

## Flat
```javascript
let convert = require('xls-to-pjson');

// Last parameter is the number pertaining to the header row in the excel file if it exists.
convert.read('input.xls', 'output.json', 1);
```

## workflow
```javascript
let convert = require('xls-to-pjson');

// Last parameter is the number pertaining to the header row in the excel file if it exists.
convert.workflow('input.xls', 'output.json', 1);
```


## CLI
Use `-r` for the header row (see the javascript version for more info).
`node cli.js -i input.xls -o output.json -r 1`

If you want to create a workflow just supply the option `--workflow`
