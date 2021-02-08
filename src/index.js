/* eslint-env browser */
/* global jexcel, bootstrap, config */

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

var projects = {};
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
    onafterchanges: function() {
        if (currentylLoading) return;
        if (saveToLocalStorageTimeout) {
            clearTimeout(saveToLocalStorageTimeout);
        }
        saveToLocalStorageTimeout = setTimeout(saveToLocalStorage, 100);
    },
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
                if (projects[c]) {
                    let project = projects[c];
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
    contextMenu: function(obj, x, y) {
        var items = [];

        if (y == null) {
            items.push({
                title: "Add blank project column",
                onclick: addNewProjectColumn,
            });

            // Rename column
            if (y >= FIRST_PROJECT_COL && obj.options.allowRenameColumn == true) {
                items.push({
                    title: obj.options.text.renameThisColumn,
                    onclick: function() {
                        obj.setHeader(x);
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
    },
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
var saveToLocalStorageTimeout = null;
const LOCAL_STROAGE_KEY = 'bom_table_data';
const channel = new BroadcastChannel('BOMBib-channel');
channel.onmessage = function (evt) {
    if (evt.data.action == 'updatedLocalStorage') {
        loadFromLocalStorage();
    }
};
function saveToLocalStorage() {
    if (currentylLoading) return;
    saveToLocalStorageTimeout = null;
    let tabledata = jexceltable.getData().slice(0, -SPARE_ROWS);
    let data = {
        "projects": projects,
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
            deleteLastColumns(tableColumnCount - importColumnCount)
        } else {
            while (jexceltable.options.data[0].length < importColumnCount) {
                addNewProjectColumn(null);
            }
        }
        projects = data.projects;
        jexceltable.setData(tabledata);
        // Use setValue to force update of component counts
        if (importColumnCount > FIRST_PROJECT_COL) {
            jexceltable.setValueFromCoords(FIRST_PROJECT_COL, PROJECT_COUNT_ROW, jexceltable.options.data[PROJECT_COUNT_ROW][FIRST_PROJECT_COL]);
        }
    }
    currentylLoading = false;
}
loadFromLocalStorage();

//FROM https://stackoverflow.com/a/35970894/2256700
var getJSON = function(url, callback) {
    var xhr = new XMLHttpRequest();
    xhr.open('GET', url, true);
    xhr.responseType = 'json';
    xhr.onload = function() {
        var status = xhr.status;
        if (status === 200) {
            callback(null, xhr.response);
        } else {
            callback(status, xhr.response);
        }
    };
    xhr.send();
};



var initializedAddProjectModal = false;

function initializeBibtab() {
    if (!initializedAddProjectModal) {
        initializedAddProjectModal = true;
        getJSON(config.librarypath, load_library);
        let tagdiv = document.getElementById('addProjectModalTagsDiv');
        let nodes = document.createDocumentFragment();
        config.projecttags.map(function (tag) {
            let node = document.createElement('div');
            node.className = 'form-check d-inline-block me-3 d-lg-block';
            node.innerHTML = '<input class="form-check-input" type="checkbox" value="" id="addProjectModalTagCheck' + tag + '"><label class="form-check-label" for="addProjectModalTagCheck' + tag + '"> ' + tag + '</label>';
            nodes.appendChild(node);
        });
        tagdiv.replaceChildren(nodes);
    }
}
document.getElementById('bib-tab').addEventListener('show.bs.tab', initializeBibtab);
function switchTab(tab) {
    if (tab == '#intro') {
        bootstrap.Tab.getInstance(document.querySelector('#tabs a[data-bs-target="#intro-tab-pane"]')).show();
    } else if (tab == '#bib') {
        bootstrap.Tab.getInstance(document.querySelector('#tabs a[data-bs-target="#bib-tab-pane"]')).show();
    } else if (tab == '#boms') {
        bootstrap.Tab.getInstance(document.querySelector('#tabs a[data-bs-target="#bom-tab-pane"]')).show();
    }
}
window.addEventListener("load", function () {
    if (window.location.hash) {
        switchTab(window.location.hash);
    }
});


var tagRegex = RegExp('\\[(' + config.projecttags.join('|') + ')\\]', 'g');

var library = [];
function parse_library_entry(p) {
    let project = {};
    project.title = p.t;
    project.author = {
        'name': p.a,
    };
    project.projectpath = p.p;
    return parse_tags(project);
}
function parse_tags(project) {
    project.tags = new Map();
    for (const match of project.title.matchAll(tagRegex)) {
        project.tags.set(match[1], true);
    }
    project.title = project.title.replace(tagRegex, '');
    return project;
}
if (!('content' in document.createElement('template'))) {
    alert("Your browser does not support the necessary feature, Sorry.");
}

function load_library(err, data) {
    let listGroup = document.getElementById('projectListGroup');
    if (err) {
        listGroup.replaceChildren('<div class="list-group-item">An error occured</div>');
        return;
    }
    let baseLibraryPath = config.librarypath.substring(0, config.librarypath.lastIndexOf('/') + 1);
    let nodes = document.createDocumentFragment();
    let template = document.getElementById('projectListGroupItemTemplate');
    let aNode = template.content.querySelector('a.list-group-item.list-group-item-action');
    let titleNode = template.content.querySelector('.projecttitle');
    let authorNode = template.content.querySelector('.projectauthor');
    let tagsNode = template.content.querySelector('.projecttags');
    data.forEach(function (p) {
        let project = parse_library_entry(p);
        library.push(project);

        aNode.href = "#project:" + baseLibraryPath + project.projectpath;
        titleNode.innerText = project.title;
        authorNode.innerText = project.author ? project.author.name : '';
        tagsNode.innerHTML = Array.from(project.tags.keys(), function (tag) {
            return '<span class="badge bg-info text-dark me-1">' + tag + '</span>';
        }).join('');
        nodes.appendChild(document.importNode(template.content, true));
    });
    listGroup.replaceChildren(nodes);
}

function escapeHTML(text) {
    // DO NOT USE FOR HTML-PARAMETERS like <a href="".
    // THIS DOES NOT ESCAPE QUOTATION MARKS
    let div = document.createElement('div');
    div.appendChild(document.createTextNode(text));
    return div.innerHTML;
}

var projectModalElement = document.getElementById('projectModal');
var projectModal = new bootstrap.Modal(projectModalElement, {'backdrop': 'static', 'keyboard': false});
var projectModalAuthorPopover = new bootstrap.Popover(projectModalElement.querySelector('.projectauthor'));
var projectModalCommitterPopover = new bootstrap.Popover(projectModalElement.querySelector('.projectcommitter'));
var currentlyLoadedProject = null;
function loadProject(project, projectpath) {
    project = parse_tags(project);
    project.projectpath = projectpath;
    //TODO Validate Project Schema
    currentlyLoadedProject = project;
    projectModalElement.querySelectorAll('[role=status]').forEach(e => e.style.display = 'none');
    projectModalElement.querySelector('#addProjectToBomButton').disabled = false;
    projectModalElement.querySelector('.modal-title').innerText = project.title;
    let projectAuthorNode = projectModalElement.querySelector('.projectauthor');
    makePersonPopover(projectAuthorNode, project.author);

    let projectCommitterNode = projectModalElement.querySelector('.projectcommitter');
    makePersonPopover(projectCommitterNode, project.committer);

    let tagsHTML =  Array.from(project.tags.keys(), function (tag) {
        return '<span class="badge bg-info text-dark me-1">' + tag + '</span>';
    }).join('');
    let projectdescriptionDiv = projectModalElement.querySelector('.projectdescription');
    projectdescriptionDiv.innerHTML = '<p>Tags: ' + tagsHTML + '</p>';
    if (project.links) {
        let youtubeembed = true;
        let linkList = document.createElement('ul');
        for (let label in project.links) {
            let li = document.createElement('li');
            let el = document.createElement('a');
            el.href = project.links[label];
            el.target = "_blank";
            el.innerText = label;
            li.appendChild(el);
            if (youtubeembed && project.links[label].substr(0, 32) == 'https://www.youtube.com/watch?v=') {
                youtubeembed = false;

                li.appendChild(document.createElement('br'));
                let divel = document.createElement('div');
                divel.className = 'ratio ratio-16x9';

                el = document.createElement('iframe');
                el.width = 560;
                el.height = 315;
                el.frameBorder = "0";
                el.allowFullscreen = true;
                el.src = project.links[label].replace('https://www.youtube.com/watch?v=', 'https://www.youtube-nocookie.com/embed/');
                el.allow = "accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture";
                divel.appendChild(el);
                li.appendChild(divel);
            }
            linkList.appendChild(li);
        }
        projectdescriptionDiv.appendChild(linkList);
    }

    projectdescriptionDiv.appendChild(document.createTextNode(project.description));

    let projectModalBOM = projectModalElement.querySelector('.projectbom tbody');
    let bomFragment = document.createDocumentFragment();
    for (const bom of project.bom) {
        let row = document.createElement('tr');
        let el = document.createElement('td');
        el.innerText = bom.type || '';
        if (bom.note) {
            el.rowSpan = 2;
        }
        row.appendChild(el);

        el = document.createElement('td');
        el.innerText = bom.value || '';
        row.appendChild(el);

        el = document.createElement('td');
        el.innerText = bom.spec || '';
        row.appendChild(el);

        el = document.createElement('td');
        el.innerText = bom.qty;
        row.appendChild(el);
        bomFragment.appendChild(row);

        if (bom.note) {
            row = document.createElement('tr');
            el = document.createElement('td');
            el.colSpan = 3;
            el.className = 'text-muted';
            el.innerText = bom.note;
            el.innerHTML = '<i class="bi bi-chat-right-text-fill"></i> ' + el.innerHTML;
            row.appendChild(el);
            bomFragment.appendChild(row);
        }
    }
    projectModalBOM.replaceChildren(bomFragment);

}

function makePersonPopover(popoverNode, person) {
    popoverNode.innerText = person.name;
    popoverNode.dataset.bsOriginalTitle = escapeHTML(person.name);
    let content = document.createDocumentFragment();
    //Create Elemenents from scratch, so we can get HTML-Parameter escaping by setting el.href
    if (person.github) {
        let el = document.createElement('A');
        el.href = 'https://github.com/' + person.github;
        el.target = "_blank";
        el.className = 'btn btn-sm btn-outline-dark m-1';
        el.innerHTML = '<i class="bi bi-github"></i> GitHub';
        content.appendChild(el);
    }
    if (person.youtube) {
        let el = document.createElement('A');
        el.href = person.youtube;
        el.target = "_blank";
        el.className = 'btn btn-sm btn-outline-dark m-1';
        el.innerHTML = '<i class="bi bi-youtube"></i> YouTube';
        content.appendChild(el);
    }
    if (person.patreon) {
        let el = document.createElement('A');
        el.href = person.patreon;
        el.target = "_blank";
        el.className = 'btn btn-sm btn-outline-dark m-1';
        el.innerHTML = 'Patreon';
        content.appendChild(el);
    }
    let div = document.createElement('div');
    div.appendChild(content);

    popoverNode.dataset.bsContent = div.innerHTML;
}

var triggerTabList = [].slice.call(document.querySelectorAll('#tabs a'));
triggerTabList.forEach(function (triggerEl) {
    var tabTrigger = new bootstrap.Tab(triggerEl);

    triggerEl.addEventListener('click', function () {
        tabTrigger.show();
    });
});

function loadProjectFromHash() {
    if (window.location.hash.substr(0, 11) == '#project:./' || window.location.hash.substr(0, 12) == '#project:../') {
        projectModalElement.querySelectorAll('[role=status]').forEach(e => e.style.display = null);
        projectModalElement.querySelector('#addProjectToBomButton').disabled = true;
        projectModal.show();
        var projectpath = window.location.hash.substr(9);
        getJSON(projectpath, function (err, data) {
            if (err) {
                alert("Could not load data");
                return;
            }
            loadProject(data, projectpath);
        });
    } else {
        projectModal.hide();
        currentlyLoadedProject = null;
        switchTab(window.location.hash);
    }
}
window.addEventListener("hashchange", loadProjectFromHash, false);
loadProjectFromHash();
projectModalElement.addEventListener('hidden.bs.modal', function () {
    if (document.getElementById('intro-tab-pane').classList.contains('active')) {
        window.location.hash = '#intro';
    } else if (document.getElementById('bib-tab-pane').classList.contains('active')) {
        window.location.hash = '#bib';
    } else if (document.getElementById('bom-tab-pane').classList.contains('active')) {
        window.location.hash = '#boms';
    } else {
        window.location.hash = '';
    }
});

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
    projects[newColumn] = project;
    return newColumn;
}


document.getElementById('addProjectToBomButton').addEventListener('click', function () {
    if (currentlyLoadedProject) {
        var bomIndex = {};

        jexceltable.options.data.forEach(function (row, i) {
            if (!row[0] && !row[1] && !row[2]) return;
            bomIndex[row[0] + '_' + row[1] + '_' + row[2]] = i;
        });

        let newColumn = addNewProjectColumn(currentlyLoadedProject);
        currentlyLoadedProject.bom.forEach((item) => {
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
        projectModal.hide();
        //let triggerEl = document.querySelector('#tabs a[href="#bom-tab-pane"]');
        //#triggerEl.click();
    }
});

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
    projects = {};
    if (jexceltable.options.data[0].length > FIRST_PROJECT_COL) {
        deleteLastColumns(jexceltable.options.data[0].length - FIRST_PROJECT_COL);
    }
    jexceltable.deleteRow(LAST_BLOCKED_ROW + 1, jexceltable.options.data.length - LAST_BLOCKED_ROW);
    saveToLocalStorage();
}
