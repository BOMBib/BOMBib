/* eslint-env browser */
/* global jexcel */

var data = [
    ['', '', '', '', '', '', 1, 2],
    ['Resistor', '100k', '1%', '0.005', '', '', 3, 2],
    ['Capacitor', '2nF', 'ceramic', '', '', '', 1, 1],
    ['Potentiometer', '100k', 'Stereo Linear', '1.5', '', '', 2, 0],
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
};

/* exported SUMCOL */
var SUMCOL = function(instance, columnId) {
    var total = 0;
    for (var j = 0; j < instance.options.data.length; j++) {
        if (parseFloat(instance.records[j][columnId - 1].innerHTML, 10)) {
            total += parseFloat(instance.records[j][columnId - 1].innerHTML, 10);
        }
    }
    return total;
};
/* exported SUMROWMUL */
var SUMROWMUL = function(instance, rowId, rowId2, startCol) {
    var total = 0;
    for (var j = startCol; j < instance.records[rowId].length; j++) {
        if (Number(instance.getValueFromCoords(j, rowId)) && Number(instance.getValueFromCoords(j, rowId2))) {
            total += Number(instance.getValueFromCoords(j, rowId)) * Number(instance.getValueFromCoords(j, rowId2));
        }
    }
    return total;
};
/* exported SUMCOLMUL */
var SUMCOLMUL = function(instance, colId, colId2, startRow) {
    var total = 0;
    for (var j = startRow; j < instance.records.length; j++) {
        if (parseFloat(instance.getValueFromCoords(colId, j), 10) && parseFloat(instance.getValueFromCoords(colId2, j), 10)) {
            total += parseFloat(instance.getValueFromCoords(colId, j), 10) * parseFloat(instance.getValueFromCoords(colId2, j), 10);
        }
    }
    return total.toFixed(3);
};

var SPARE_COLUMNS = 3;
var PER_PART_COST_COL = 3;
var PART_COUNT_COLUMN = 4;
var PER_PART_COST_SUM_COL = 5;
var FIRST_PROJECT_COL = 6;

function addDependency(instance, cell, dependson) {
    if (dependson in instance.jexcel.formula) {
        instance.jexcel.formula[dependson].push(cell);
    } else {
        instance.jexcel.formula[dependson] = [cell];
    }
}

var project_count = 2;
/* exported jexceltable */
var jexceltable = jexcel(document.getElementById('spreadsheet'), {
    data: data,
    minSpareRows: SPARE_COLUMNS,
    columnSorting: false,
    allowInsertColumn: false, // TODO: This needs to be enabled to add new columns later. Instead remove option from context menu.
    allowManualInsertColumn: false,
    allowDeleteColumn: false,
    columns: [
        { type: 'dropdown', title: 'Part Category', width: 120, source: part_categorys },
        { type: 'text', title: 'Value', width: 100,  },
        //TODO: we want a text-autocomplete, not a drop-autocomplete
        //{ type: 'autocomplete', title:'Specification', width:100,  source: spec_all_autocompletes, multiple:true, filter: dropdownFilter},
        { type: 'text', title: 'Specification', width: 100, },
        { type: 'numerical', title: 'Cost per Part', width: 100, },
        { type: 'numerical', title: 'Count', width: 100, },
        { type: 'numerical', title: 'Cost', width: 100, },

        { type: 'numerical', title: 'Project 1', width: 80 },
        { type: 'numerical', title: 'Project 2', width: 80 },
    ],
    footers: [[
        '',
        '',
        '',
        'Total',
        '=SUMCOL(TABLE(), COLUMN())', '=SUMCOL(TABLE(), COLUMN()) + "¤"',
        '=VALUE(COLUMN(), 1) + SUMCOLMUL(TABLE(), COLUMN() - 1, ' + PER_PART_COST_COL + ', 1) + "¤"',
        '=VALUE(COLUMN(), 1) + SUMCOLMUL(TABLE(), COLUMN() - 1, ' + PER_PART_COST_COL + ', 1) + "¤"',
    ]],
    updateTable: function(instance, cell, c, r, source, value, id) {
        if (r == 0 && c < FIRST_PROJECT_COL) {
            cell.classList.add('readonly');
            cell.classList.remove('jexcel_dropdown');
            cell.innerHTML = c == FIRST_PROJECT_COL - 1 ? 'Count:' : '';
            cell.style.backgroundColor = '#f3f3f3';
        }
        if (r == 0 && c >= FIRST_PROJECT_COL) {
            if (cell.innerText > 0) {
                cell.innerHTML = cell.innerText + 'x';
            }
        }
        if (r > 0 && c == PER_PART_COST_COL && value) {
            cell.innerHTML = parseFloat(instance.jexcel.options.data[r][c]).toFixed(3) + '¤';
        }

        if (r > 0 && c == PER_PART_COST_SUM_COL) {
            if (!instance.jexcel.getValue(id) && instance.jexcel.rows.length - r > SPARE_COLUMNS) {
                instance.jexcel.setValue(id, '=IF(D' + (r + 1) + ', D' + (r + 1) + ' * E' + (r + 1) + ", '')", true);
            }
            if (cell.innerHTML && cell.innerHTML[cell.innerHTML.length - 1] != '¤') {
                cell.innerHTML = parseFloat(value).toFixed(3) + '¤';
            }
            cell.classList.add('readonly');
        }
        if (r > 0 && c == PART_COUNT_COLUMN) {
            let part_count_formula = '=SUMROWMUL(TABLE(), ROW() - 1, 0, ' + FIRST_PROJECT_COL + ')';
            if (instance.jexcel.getValueFromCoords(c, r) != part_count_formula) {
                if (instance.jexcel.rows.length - r > SPARE_COLUMNS) {
                    instance.jexcel.setValue(id, part_count_formula, true);
                    //We have to set the dependencies of this field manually, because jexcel fails to figure them out.
                    for (let j = 0; j < project_count; j++) {
                        let cellName = jexcel.getColumnName(PART_COUNT_COLUMN) + (r + 1);
                        addDependency(instance, cellName, jexcel.getColumnName(FIRST_PROJECT_COL + j) + (r + 1));
                        addDependency(instance, cellName, jexcel.getColumnName(FIRST_PROJECT_COL + j) + "1");
                    }
                }
            }
            cell.classList.add('readonly');
        }
    },
});


jexceltable.setComments(jexcel.getColumnName(FIRST_PROJECT_COL + 1) + '2', 'This is a comment for the resistor in Project 2');
