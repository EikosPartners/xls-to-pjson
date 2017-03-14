let argv = require('minimist')(process.argv.slice(2));
let convert = require('./convert');

if (argv.h) {
    console.log('Options: ');
    console.log('           -h : Help, display this list');
    console.log('           -i : Input filename (Excel sheet)');
    console.log('           -o : Output filename (JSON)');
    console.log('           -r : The row corresponding to the header in the excel sheet if it exists');
    console.log('   --workflow : Supply this option if you want to create a workflow instead of a flat structure');
    console.log('     --series : Supply this option if you want to create a series of actions');
    process.exit();
}

if (!argv.i || !argv.o) {
    console.error("You must provide the input and output file names with -i -o");
    process.exit();
}

let series = argv.series ? true : false;

if (argv.workflow) {
    convert.workflow(argv.i, argv.o, series, argv.r);
} else {
    convert.flat(argv.i, argv.o, series, argv.r);
}
