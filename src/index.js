/* eslint-env browser */
/* global jexcel, bootstrap, config */

var data = [
    ['', '', '', '', '', ''],
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

var SPARE_COLUMNS = 3;
var PER_PART_COST_COL = 3;
var PART_COUNT_COLUMN = 4;
var PER_PART_COST_SUM_COL = 5;
var FIRST_PROJECT_COL = 6;

const PROJECT_FOOTER_FORMULA = '=VALUE(COLUMN(), 1) + SUMCOLMUL(TABLE(), COLUMN() - 1, ' + PER_PART_COST_COL + ', 1) + "¤"';

function addDependency(instance, cell, dependson) {
    if (dependson in instance.formula) {
        instance.formula[dependson].push(cell);
    } else {
        instance.formula[dependson] = [cell];
    }
}

var project_count = 0;
/* exported jexceltable */
var jexceltable = jexcel(document.getElementById('spreadsheet'), {
    data: data,
    minSpareRows: SPARE_COLUMNS,
    columnSorting: false,
    allowManualInsertColumn: false,
    allowDeleteColumn: false,
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
                        addDependency(instance.jexcel, cellName, jexcel.getColumnName(FIRST_PROJECT_COL + j) + (r + 1));
                        addDependency(instance.jexcel, cellName, jexcel.getColumnName(FIRST_PROJECT_COL + j) + "1");
                    }
                }
            }
            cell.classList.add('readonly');
        }
    },
});


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
window.addEventListener("load", function () {
    if (window.location.hash) {
        if (window.location.hash == '#bib') {
            bootstrap.Tab.getInstance(document.querySelector('#tabs a[data-bs-target="#bib-tab-pane"]')).show();
        } else if (window.location.hash == '#boms') {
            bootstrap.Tab.getInstance(document.querySelector('#tabs a[data-bs-target="#bom-tab-pane"]')).show();
        }
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
function loadProject(project) {
    project = parse_tags(project);
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
        getJSON(window.location.hash.substr(9), function (err, data) {
            if (err) {
                alert("Could not load data");
                return;
            }
            loadProject(data);
        });
    } else {
        projectModal.hide();
        currentlyLoadedProject = null;
    }
}
window.addEventListener("hashchange", loadProjectFromHash, false);
loadProjectFromHash();
projectModalElement.addEventListener('hidden.bs.modal', function () {
    window.location.hash = '';
});

function escapeCellContent(content) {
    if (content && content[0] == '=') {
        return "=" + JSON.stringify(content);
    } else {
        return content;
    }
}

function addNewProjectColumn() {
    let newColumn = jexceltable.colgroup.length;
    project_count += 1;
    jexceltable.insertColumn(1, newColumn, false, { type: 'numerical', title: 'Project ' + (jexceltable.colgroup.length - FIRST_PROJECT_COL + 1), width: 80 });
    jexceltable.setValueFromCoords(newColumn, 0, 1); //Set count for new Project
    jexceltable.options.footers[0][newColumn] = PROJECT_FOOTER_FORMULA;
    for (let r = 0; r < jexceltable.rows.length - SPARE_COLUMNS; r++) {
        let cellName = jexcel.getColumnName(PART_COUNT_COLUMN) + (r + 1);
        addDependency(jexceltable, cellName, jexcel.getColumnName(newColumn) + (r + 1));
        addDependency(jexceltable, cellName, jexcel.getColumnName(newColumn) + "1");
    }
    return newColumn;
}


document.getElementById('addProjectToBomButton').addEventListener('click', function () {
    if (currentlyLoadedProject) {
        var bomIndex = {};

        jexceltable.options.data.forEach(function (row, i) {
            if (!row[0] && !row[1] && !row[2]) return;
            bomIndex[row[0] + '_' + row[1] + '_' + row[2]] = i;
        });

        let newColumn = addNewProjectColumn();
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
                rownum = jexceltable.rows.length - SPARE_COLUMNS - 1;
                jexceltable.insertRow(row, rownum);
                rownum = rownum + 1;
                let cellName = jexcel.getColumnName(PART_COUNT_COLUMN) + (rownum + 1);
                addDependency(jexceltable, cellName, jexcel.getColumnName(newColumn) + (rownum + 1));
                addDependency(jexceltable, cellName, jexcel.getColumnName(newColumn) + "1");
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
    for (let i = 1; i < jexceltable.rows.length - SPARE_COLUMNS; i++) {
        rowIndex.push(i);
    }
    let tabledata = jexceltable.options.data;
    let compString = (a, b) =>  a && b ? a.localeCompare(b) : (a ? -1 : (b ? 1 : 0));
    const valueRegex = /^\s*(\d+(\.\d*)?)\s*([kK]|M|u|n)?(F)?\s*$/;
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
        jexceltable.moveRow(oldIndex, newIndex + 1);
    });

}
