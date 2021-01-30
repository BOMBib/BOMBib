/*jslint browser:true, devel: true */

var data = [
    ['','','','','','',1,2],
    ['Resistor', '100k', '1%', '0.005','','',3,2],
    ['Capacitor', '2nF', 'ceramic', '','','',1,1],
    ['Potentiometer', '100k', 'Stereo Linear', '1.5','','',2,0],
];

var part_categorys = [
    'Resistor', 'Capacitor', 'Potentiometer', 'Diode', 'IC', ''
];

var spec_autocompletes = {
    'Potentiometer': ['Mono', 'Stereo', 'Linear', 'Logarithmic'],
    'Resistor': ['Matched']
};
var spec_all_autocompletes = Object.values(spec_autocompletes).flat();

var dropdownFilter = function (instance, cell, c, r, source) {
    var value = instance.jexcel.getValueFromCoords(0, r);
    return spec_autocompletes[value] || source;
}

var SUMCOL = function(instance, columnId) {
    var total = 0;
    for (var j = 0; j < instance.options.data.length; j++) {
        if (Number(instance.records[j][columnId-1].innerHTML)) {
            total += Number(instance.records[j][columnId-1].innerHTML);
        }
    }
    return total;
}
var SUMROWMUL = function(instance, rowId, rowId2, startCol) {
    var total = 0;
    for (var j = startCol; j < instance.records[rowId].length; j++) {
        if (Number(instance.records[rowId][j].innerHTML) && Number(instance.records[rowId2][j].innerHTML)) {
            total += Number(instance.records[rowId][j].innerHTML) * Number(instance.records[rowId2][j].innerHTML);
        }
    }
    return total;
}
var SUMPRODUCT = function(a,b) {
    var total = 0;
    for (var j in a) {
        total += a[j] * b[j]
    }
    return total;
}
var COL_TO_LETTER = ['A', 'B', 'C', 'D', 'E']
var SPARE_COLUMNS = 3;
var PART_COUNT_COLUMN = 4;
var PER_PART_COST_SUM_COL = 5;
var FIRST_PROJECT_COL = 6;

var project_count = 2;
var jexceltable = jexcel(document.getElementById('spreadsheet'), {
    data:data,
    minSpareRows: SPARE_COLUMNS,
    columnSorting:false,
    allowInsertColumn: false, // TODO: This needs to be enabled to add new columns later. Instead remove option from context menu.
    allowManualInsertColumn: false,
    allowDeleteColumn: false,
    columns: [
        { type: 'dropdown', title:'Part Category', width:120, source: part_categorys },
        { type: 'text', title:'Value', width:100,  },
        //TODO: we want a text-autocomplete, not a drop-autocomplete
        //{ type: 'autocomplete', title:'Specification', width:100,  source: spec_all_autocompletes, multiple:true, filter: dropdownFilter},
        { type: 'text', title:'Specification', width: 100, },
        { type: 'numerical', title: 'Cost per Part', width: 100, },
        { type: 'numerical', title: 'Count', width: 100, },
        { type: 'numerical', title: 'Cost', width: 100, },

        { type: 'numerical', title:'Project 1', width:80 },
        { type: 'numerical', title:'Project 2', width:80 },
    ],
    footers: [['','', '', 'Total','=SUMCOL(TABLE(), COLUMN())','=SUMCOL(TABLE(), COLUMN())']],
    updateTable: function(instance, cell, c, r, source, value, id) {
        if (r == 0 && c < FIRST_PROJECT_COL) {
            cell.classList.add('readonly');
            cell.classList.remove('jexcel_dropdown');
            cell.innerHTML = c == FIRST_PROJECT_COL - 1 ? 'Count:' : '';
            cell.style.backgroundColor = '#f3f3f3';
        }
        if (r== 0 && c>= FIRST_PROJECT_COL) {
            cell.innerHTML
        }
        if (r > 0 && c == PER_PART_COST_SUM_COL  && !instance.jexcel.getValue(id)) {
            if (instance.jexcel.rows.length - r > SPARE_COLUMNS) {
                instance.jexcel.setValue(id, '=D' + (r+1) + ' * E' + (r+1), true);
            }
            cell.style.backgroundColor = '#f3f3f3';
            //cell.classList.add('readonly');
        }
        if (r > 0 && c == PART_COUNT_COLUMN && !value) {
            if (instance.jexcel.rows.length - r > SPARE_COLUMNS) {
                instance.jexcel.setValue(id, '=SUMROWMUL(TABLE(), ROW() - 1, 0, ' + FIRST_PROJECT_COL + ')', true);
                //We have to set the dependencies of this field manually, because jexcel fails to figure them out.
                for (let j = 0; j < project_count; j++) {
                    let cellName = jexcel.getColumnName(PART_COUNT_COLUMN) + (r+1);
                    instance.jexcel.formula[jexcel.getColumnName(FIRST_PROJECT_COL + j) + (r+1)] = [cellName]
                    instance.jexcel.formula[jexcel.getColumnName(FIRST_PROJECT_COL + j) + "1"] = [cellName]
                }
            }
            cell.style.backgroundColor = '#f3f3f3';
            //cell.classList.add('readonly');
        }
    },
});
