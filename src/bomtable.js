/* eslint-env browser */
/* global jexcel, config */

var data = [
    ['', '', '', '', '', ''],
    ['x', '', '', '', '', ''],
];

let part_categorys = [''].concat(config.categorys);

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

const LAST_BLOCKED_ROW = 1;
const SPARE_ROWS = 3;
const PER_PART_COST_COL = 3;
const PART_COUNT_COLUMN = 4;
const PER_PART_COST_SUM_COL = 5;
const FIRST_PROJECT_COL = 6;

const PROJECT_TITLE_ROW = 0;
const PROJECT_COUNT_ROW = 1;

const PROJECT_FOOTER_FORMULA = '=VALUE(COLUMN(), ' + (PROJECT_COUNT_ROW + 1) + ') + SUMCOLMUL(TABLE(), COLUMN() - 1, ' + PER_PART_COST_COL + ', 1) + "¤"';

var projectsInBOMTable = {};

function queueSave() {
    if (currentylLoading) return;
    /* global debounceUtil */
    debounceUtil('saveBomTableToLocalStorage', 200, 1000, saveToLocalStorage);
}

/* exported jexceltable */
var jexceltable = jexcel(document.getElementById('spreadsheet'), {
    data: data,
    minSpareRows: SPARE_ROWS,
    columnSorting: false,
    allowManualInsertColumn: false,
    columns: [
        { type: 'dropdown', title: 'Part Category', width: 80, source: part_categorys },
        { type: 'text', title: 'Value', width: 80 },
        //TODO: we want a text-autocomplete, not a drop-autocomplete
        //{ type: 'autocomplete', title:'Specification', width:100,  source: spec_all_autocompletes, multiple:true, filter: dropdownFilter},
        { type: 'text', title: 'Specification', width: 200 },
        { type: 'numerical', title: 'Cost per Part', width: 100 },
        { type: 'numerical', title: 'Count', width: 100 },
        { type: 'numerical', title: 'Cost', width: 100 },
    ],
    footers: [[
        '',
        '',
        '',
        'Total',
        '=SUMCOL(TABLE(), COLUMN())', '=SUMCOL(TABLE(), COLUMN()) + "¤"',
    ]],
    onmoverow: queueSave,
    onafterchanges: queueSave,
    updateTable: function(instance, cell, c, r, source, value, id) {
        if (r == PROJECT_TITLE_ROW) {
            if (c < FIRST_PROJECT_COL) {
                cell.classList.add('readonly');
                cell.classList.remove('jexcel_dropdown');
                if (c == FIRST_PROJECT_COL - 1) {
                    cell.innerHTML = 'Title:';
                }
            } else {
                cell.classList.add('text-wrap');
                if (projectsInBOMTable[c]) {
                    let project = projectsInBOMTable[c];
                    cell.classList.add('readonly');
                    cell.innerHTML = '<a href="#project:' + project.projectpath + '">' + project.title + '</a>';
                }
            }
        }
        if (r == PROJECT_COUNT_ROW && c < FIRST_PROJECT_COL) {
            cell.classList.add('readonly');
            cell.classList.remove('jexcel_dropdown');
            cell.innerHTML = c == FIRST_PROJECT_COL - 1 ? 'Count:' : '';
        }
        if (r == PROJECT_COUNT_ROW && c >= FIRST_PROJECT_COL) {
            if (instance.jexcel.options.data[r][c] != '' && Number.isFinite(Number(instance.jexcel.options.data[r][c]))) {
                cell.innerHTML = (instance.jexcel.options.data[r][c]) + 'x';
            } else {
                cell.innerHTML = '';
            }
        }
        if (r > LAST_BLOCKED_ROW && c == PER_PART_COST_COL && value) {
            cell.innerHTML = parseFloat(instance.jexcel.options.data[r][c]).toFixed(3) + '¤';
        }

        if (r > LAST_BLOCKED_ROW && c == PER_PART_COST_SUM_COL) {
            if (!instance.jexcel.getValue(id) && instance.jexcel.rows.length - r > SPARE_ROWS) {
                instance.jexcel.setValue(id, '=IF(D' + (r + 1) + ', D' + (r + 1) + ' * E' + (r + 1) + ", '')", true);
            }
            if (cell.innerHTML && cell.innerHTML[cell.innerHTML.length - 1] != '¤') {
                cell.innerHTML = parseFloat(value).toFixed(3) + '¤';
            }
            cell.classList.add('readonly');
        }
        if (r > LAST_BLOCKED_ROW && c == PART_COUNT_COLUMN) {
            let part_count_formula = '=SUMROWMUL(TABLE(), ROW() - 1, ' + PROJECT_COUNT_ROW + ', ' + FIRST_PROJECT_COL + ')';
            if (cell.innerHTML == '0') {
                cell.parentNode.classList.add('unused_part_row');
            } else {
                cell.parentNode.classList.remove('unused_part_row');
            }
            if (instance.jexcel.getValueFromCoords(c, r) != part_count_formula) {
                if (instance.jexcel.rows.length - r > SPARE_ROWS) {
                    instance.jexcel.setValue(id, part_count_formula, true);
                }
            }
            cell.classList.add('readonly');
        }
        setDependencies(instance.jexcel, c, r);
    },
    contextMenu: bomtablecontextmenu,
});

function setDependencies(instance, c, r) {
    let cellName = jexcel.getColumnName(c) + (r + 1);
    if (c >= FIRST_PROJECT_COL && r > LAST_BLOCKED_ROW) {
        // When component count in project changes the total component count changes
        if (!instance.formula[cellName] || instance.formula[cellName].length != 1) {
            instance.formula[cellName] = [jexcel.getColumnName(PART_COUNT_COLUMN) + (r + 1)];
        }
    }
    if (c >= FIRST_PROJECT_COL && r == PROJECT_COUNT_ROW) {
        //When a project count changes all component counts change
        if (!instance.formula[cellName] || instance.formula[cellName].length != instance.options.data.length - LAST_BLOCKED_ROW - 1) {
            let deps = [];
            for (let i = LAST_BLOCKED_ROW + 1; i < instance.options.data.length; i++) {
                deps.push(jexcel.getColumnName(PART_COUNT_COLUMN) + (i + 1));
            }
            instance.formula[cellName] = deps;
        }
    }

}

var currentylLoading = true;

const LOCAL_STROAGE_KEY = 'bom_table_data';
const channel = new BroadcastChannel('BOMBib-channel');
channel.onmessage = function (evt) {
    if (evt.data.action == 'updatedLocalStorage') {
        loadFromLocalStorage();
    }
};
function saveToLocalStorage() {
    if (currentylLoading) return;
    let tabledata = jexceltable.getData().slice(0, -SPARE_ROWS);
    let data = {
        "projects": projectsInBOMTable,
        "tabledata": tabledata,
    };
    localStorage.setItem(LOCAL_STROAGE_KEY, JSON.stringify(data));
    channel.postMessage({
        'action': 'updatedLocalStorage',
    });
}

function loadFromLocalStorage() {
    let data = localStorage.getItem(LOCAL_STROAGE_KEY);
    if (data) {
        currentylLoading = true;
        data = JSON.parse(data);
        let tabledata = data.tabledata;

        let importColumnCount = tabledata[0].length;
        let tableColumnCount = jexceltable.options.data[0].length;
        if (tableColumnCount > importColumnCount) {
            deleteLastColumns(tableColumnCount - importColumnCount);
        } else {
            while (jexceltable.options.data[0].length < importColumnCount) {
                addNewProjectColumn(null);
            }
        }
        projectsInBOMTable = data.projects;
        jexceltable.setData(tabledata);
        // Use setValue to force update of component counts
        if (importColumnCount > FIRST_PROJECT_COL) {
            jexceltable.setValueFromCoords(FIRST_PROJECT_COL, PROJECT_COUNT_ROW, jexceltable.options.data[PROJECT_COUNT_ROW][FIRST_PROJECT_COL]);
        }
    }
    currentylLoading = false;
}
loadFromLocalStorage();

function escapeCellContent(content) {
    if (content && content[0] == '=') {
        return "=" + JSON.stringify(content);
    } else {
        return content;
    }
}

function addNewProjectColumn(project) {
    let newColumn = jexceltable.colgroup.length;
    jexceltable.insertColumn(1, newColumn, false, { type: 'numerical', title: 'Project ' + (jexceltable.colgroup.length - FIRST_PROJECT_COL + 1), width: 80 });
    jexceltable.setValueFromCoords(newColumn, PROJECT_COUNT_ROW, 1); //Set count for new Project
    jexceltable.options.footers[0][newColumn] = PROJECT_FOOTER_FORMULA;
    projectsInBOMTable[newColumn] = project;
    return newColumn;
}


/* exported addProjectToBom */
function addProjectToBom(project, callback) {
    if (project) {
        var bomIndex = {};

        jexceltable.options.data.forEach(function (row, i) {
            if (!row[0] && !row[1] && !row[2]) return;
            bomIndex[row[0] + '_' + row[1] + '_' + row[2]] = i;
        });

        let newColumn = addNewProjectColumn(project);
        project.bom.forEach((item) => {
            let key = (item.type || '') + '_' + (item.value || '') + '_' + (item.spec || '');
            let rownum = null;
            if (key in bomIndex) {
                rownum =  bomIndex[key];
                jexceltable.setValueFromCoords(newColumn, rownum, item.qty);
            } else {
                const row = new Array(newColumn).fill('');
                row[0] = escapeCellContent(item.type);
                row[1] = escapeCellContent(item.value);
                row[2] = escapeCellContent(item.spec);
                row[newColumn] = Number(item.qty);
                rownum = jexceltable.rows.length - SPARE_ROWS - 1;
                jexceltable.insertRow(row, rownum);
                rownum = rownum + 1;
            }
            if (item.note) {
                jexceltable.setComments(jexcel.getColumnName(newColumn) + (rownum + 1), item.note);
            }
        });
        if (callback) {
            callback();
        }
        //let triggerEl = document.querySelector('#tabs a[href="#bom-tab-pane"]');
        //#triggerEl.click();
    }
}

/* exported sortBOMRows */
function sortBOMRows() {
    let rowIndex = [];
    for (let i = LAST_BLOCKED_ROW + 1; i < jexceltable.rows.length - SPARE_ROWS; i++) {
        rowIndex.push(i);
    }
    let tabledata = jexceltable.options.data;
    let compString = (a, b) =>  a && b ? a.localeCompare(b) : (a ? -1 : (b ? 1 : 0));
    const valueRegex = /^\s*(\d+(\.\d*)?)\s*(G|M|[kK]|u|n|p)?(F)?\s*$/;
    const unitMap = {
        'G': 1e9,
        'M': 1e6,
        'k': 1e3,
        'K': 1e3,
        'u': 1e-6,
        'n': 1e-9,
        'p': 1e-12,
        '': 1,
    };
    let compValue = function (a, b) {
        const matchA = a.match(valueRegex);
        if (matchA) {
            const matchB = b.match(valueRegex);
            if (matchB) {
                const valueA = parseFloat(matchA[1], 10) * unitMap[matchA[3]];
                const valueB = parseFloat(matchB[1], 10) * unitMap[matchB[3]];
                return valueB - valueA;
            }
        }
        return compString(a, b);
    };
    rowIndex.sort(function (a, b) {
        let cmp = compString(tabledata[a][0], tabledata[b][0]);
        if (cmp != 0) return cmp;
        cmp = compValue(tabledata[a][1], tabledata[b][1]);
        if (cmp != 0) return cmp;
        cmp = compString(tabledata[a][2], tabledata[b][2]);
        return cmp;
    });

    // Correct for the fact, that the index change after sorting.
    // Increment all indexes that get shiftet by moving the row i
    for (let i = 0; i < rowIndex.length; i++) {
        for (let j = i + 1 ; j < rowIndex.length; j++) {
            if (rowIndex[j] < rowIndex[i]) {
                rowIndex[j]++;
            }
        }
    }

    rowIndex.forEach((oldIndex, newIndex) => {
        jexceltable.moveRow(oldIndex, newIndex + LAST_BLOCKED_ROW + 1);
    });

}

function deleteLastColumns(n) {
    let firstToDelete = jexceltable.options.data[0].length - n;
    jexceltable.deleteColumn(firstToDelete, n);
    document.getElementById('spreadsheet').querySelector('tfoot').innerHTML = '';
    jexceltable.setFooter([jexceltable.options.footers[0].slice(0, FIRST_PROJECT_COL - 1)]);
}

/* exported clearBOMTable */
function clearBOMTable() {
    projectsInBOMTable = {};
    if (jexceltable.options.data[0].length > FIRST_PROJECT_COL) {
        deleteLastColumns(jexceltable.options.data[0].length - FIRST_PROJECT_COL);
    }
    jexceltable.deleteRow(LAST_BLOCKED_ROW + 1, jexceltable.options.data.length - LAST_BLOCKED_ROW);
    saveToLocalStorage();
}

function deleteProject(colnumber) {
    for (let i = colnumber + 1; i < jexceltable.options.data[0].length; i++) {
        projectsInBOMTable[i - 1] = projectsInBOMTable[i];
    }
    delete projectsInBOMTable[jexceltable.options.data[0].length];
    jexceltable.deleteColumn(colnumber);
    document.getElementById('spreadsheet').querySelector('tfoot').innerHTML = '';
    jexceltable.setFooter([jexceltable.options.footers[0].slice(0, -1)]);
    saveToLocalStorage();
}


function bomtablecontextmenu(obj, x, y) {
    var items = [];

    if (y == null) {
        items.push({
            title: "Add blank project column",
            onclick: addNewProjectColumn,
        });

        // Rename column
        if (x >= FIRST_PROJECT_COL && obj.options.allowRenameColumn == true) {
            items.push({
                title: obj.options.text.renameThisColumn,
                onclick: function() {
                    obj.setHeader(x);
                },
            });
        }
        if (x >= FIRST_PROJECT_COL) {
            items.push({
                title: "Remove Project",
                onclick: function () {
                    deleteProject(y);
                },
            });
        }

    } else {
        // Insert new row
        if (y >= LAST_BLOCKED_ROW) {
            if (y >= LAST_BLOCKED_ROW + 1) {
                items.push({
                    title: obj.options.text.insertANewRowBefore,
                    onclick: function() {
                        obj.insertRow(1, parseInt(y), 1);
                    },
                });
            }

            items.push({
                title: obj.options.text.insertANewRowAfter,
                onclick: function() {
                    obj.insertRow(1, parseInt(y));
                },
            });
        }

        if (y > LAST_BLOCKED_ROW && obj.options.allowDeleteRow == true) {
            items.push({
                title: obj.options.text.deleteSelectedRows,
                onclick: function() {
                    obj.deleteRow(obj.getSelectedRows().length ? undefined : parseInt(y));
                },
            });
        }

        if (x) {
            if (obj.options.allowComments == true) {
                items.push({ type: 'line' });

                var title = obj.records[y][x].getAttribute('title') || '';

                items.push({
                    title: title ? obj.options.text.editComments : obj.options.text.addComments,
                    onclick: function() {
                        var comment = prompt(obj.options.text.comments, title);
                        if (comment) {
                            obj.setComments([ x, y ], comment);
                        }
                    },
                });

                if (title) {
                    items.push({
                        title: obj.options.text.clearComments,
                        onclick: function() {
                            obj.setComments([ x, y ], '');
                        },
                    });
                }
            }
        }
    }

    // Line
    if (y >= LAST_BLOCKED_ROW) {
        items.push({ type: 'line' });
    }

    // Copy
    items.push({
        title: obj.options.text.copy,
        shortcut: 'Ctrl + C',
        onclick: function() {
            obj.copy(true);
        },
    });

    // Paste
    if (navigator && navigator.clipboard) {
        items.push({
            title: obj.options.text.paste,
            shortcut: 'Ctrl + V',
            onclick: function() {
                if (obj.selectedCell) {
                    navigator.clipboard.readText().then(function(text) {
                        if (text) {
                            jexcel.current.paste(obj.selectedCell[0], obj.selectedCell[1], text);
                        }
                    });
                }
            },
        });
    }

    // Save
    if (obj.options.allowExport) {
        items.push({
            title: obj.options.text.saveAs,
            shortcut: 'Ctrl + S',
            onclick: function () {
                obj.download();
            },
        });
    }

    // About
    if (obj.options.about) {
        items.push({
            title: obj.options.text.about,
            onclick: function() {
                alert(obj.options.about);
            },
        });
    }

    return items;
}
