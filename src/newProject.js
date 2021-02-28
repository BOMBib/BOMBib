/* eslint-env browser */
/* global bootstrap, config, jexcel */
/* global SPARE_ROWS */

/* exported newProjectModalElement, newProjectModal, */
var newProjectModalElement = document.getElementById('newProjectModal');
var newProjectModal = new bootstrap.Modal(newProjectModalElement, {'backdrop': 'static', 'keyboard': false});

const LOCALSTORAGE_LOCAL_PROJECT_PREFIX = 'local_project_meta_';
const LOCALSTORAGE_LOCAL_PROJECT_BOM_PREFIX = 'local_project_bom_';
var currentlyLoadedLocalProjectHash = null;

function queueSaveLocalProjectTable() {
    /* global debounceUtil */
    debounceUtil('saveLocalProjectToStorage', 500, 1000, function () {
        saveLocalProjectToStorage(currentlyLoadedLocalProjectHash);
    });
}

/* exported newProjectTable */
var newProjectTable = null;
function initializeNewProjectModal() {
    const NEW_PROJECT_CATEGORY_COL = 0;
    const NEW_PROJECT_VALUE_COL = 1;
    const NEW_PROJECT_SPEC_COL = 2;
    const NEW_PROJECT_NOTE_COL = 3;
    const NEW_PROJECT_QTY_COL = 4;
    let tagDiv = document.getElementById('newProjectModalTagsDiv');
    let template = document.getElementById('newProjectModalTagElementTemplate');
    let label = template.content.querySelector('label');
    let input = template.content.querySelector('input');
    let fragment = document.createDocumentFragment();
    for (let tag of config.projecttags) {
        label.innerText = tag;
        label.for = 'newProjectModalTag' + tag;
        input.id = label.for;
        input.value = tag;
        fragment.appendChild(document.importNode(template.content, true));
    }
    tagDiv.replaceChildren(fragment);

    newProjectTable = jexcel(document.getElementById('newProjectSpreadsheet'), {
        data: [[]],
        minSpareRows: SPARE_ROWS,
        columnSorting: false,
        defaultColWidth: 80,
        allowManualInsertColumn: false,
        tableOverflow: true,
        freezeColumns: 5,
        tableHeight: '600px',
        onafterchanges: queueSaveLocalProjectTable,
        onmoverow: queueSaveLocalProjectTable,
        columns: [
            { type: 'dropdown', title: 'Part Category', width: 80, source: [''].concat(config.categorys) },
            { type: 'text', title: 'Value', width: 80 },
            { type: 'text', title: 'Specification', width: 200 },
            { type: 'text', title: 'Note', width: 200 },
            { type: 'text', title: 'Quantity', width: 80 },
        ],
        updateTable: function(instance, cell, c, r, source, value, id) {
            if (c != NEW_PROJECT_CATEGORY_COL && c != NEW_PROJECT_VALUE_COL && c != NEW_PROJECT_SPEC_COL && c != NEW_PROJECT_NOTE_COL && c != NEW_PROJECT_QTY_COL) {
                cell.classList.add('non_import_part_col');
            }
            if (c == NEW_PROJECT_NOTE_COL) {
                cell.classList.add('text-wrap');
            }
            if (c == NEW_PROJECT_QTY_COL) {
                let qty = Number(cell.innerHTML);
                if (qty == 0 || isNaN(qty)) {
                    cell.parentNode.classList.add('unused_part_row');
                } else {
                    cell.parentNode.classList.remove('unused_part_row');
                }
            }
        },
    });

    newProjectModalElement.addEventListener('change', queueSaveLocalProjectTable);

    newProjectTable.insertColumn(1);
}
initializeNewProjectModal();

function saveLocalProjectToStorage(hash) {
    let project = {};

    project.title = newProjectModalElement.querySelector('#newProjectTitle').value;
    project.description = newProjectModalElement.querySelector('#newProjectDescription').value;
    function getPerson(idPrefix, element) {
        let person = {};
        person.name = element.querySelector(idPrefix + 'Name').value;
        person.github = element.querySelector(idPrefix + 'GitHub').value;
        person.youtube = element.querySelector(idPrefix + 'YouTube').value;
        person.patreon = element.querySelector(idPrefix + 'Patreon').value;
        person.web =  element.querySelector(idPrefix + 'Web').value;
        return person;
    }
    project.author = getPerson('#newProjectAuthor', newProjectModalElement);
    project.committer = getPerson('#newProjectCommitter', newProjectModalElement);
    project.tags = {};
    newProjectModalElement.querySelectorAll(
        'input[type=checkbox][name=newProjectModalTagCheckbox]:checked'
    ).forEach(function (item, i) {
        project.tags[item.value] = true;
    });

    let tabledata = newProjectTable.getData().slice(0, -SPARE_ROWS);

    localStorage.setItem(LOCALSTORAGE_LOCAL_PROJECT_BOM_PREFIX + hash, JSON.stringify(tabledata));

    localStorage.setItem(LOCALSTORAGE_LOCAL_PROJECT_PREFIX + hash, JSON.stringify(project));
}

function loadLocalProjectFromStorage(hash) {
    currentlyLoadedLocalProjectHash = hash;
    let project = JSON.parse(localStorage.getItem(LOCALSTORAGE_LOCAL_PROJECT_PREFIX + hash));

    if (!project) {
        return;
    }

    newProjectModalElement.querySelector('#newProjectTitle').value = project.title;
    newProjectModalElement.querySelector('#newProjectDescription').value = project.description;

    function loadPerson(idPrefix, element, person) {
        element.querySelector(idPrefix + 'Name').value = person.name || '';
        element.querySelector(idPrefix + 'GitHub').value = person.github || '';
        element.querySelector(idPrefix + 'YouTube').value = person.youtube || '';
        element.querySelector(idPrefix + 'Patreon').value = person.patreon || '';
        element.querySelector(idPrefix + 'Web').value = person.web || '';
    }
    loadPerson('#newProjectAuthor', newProjectModalElement, project.author || {});
    loadPerson('#newProjectCommitter', newProjectModalElement, project.committer || {});

    if (project.tags) {
        newProjectModalElement.querySelectorAll(
            'input[type=checkbox][name=newProjectModalTagCheckbox]'
        ).forEach(function (item, i) {
            item.checked = project.tags[item.value] || false;
        });
    }
    let tabledata = JSON.parse(localStorage.getItem(LOCALSTORAGE_LOCAL_PROJECT_BOM_PREFIX + hash));
    if (tabledata) {
        newProjectTable.setData(tabledata);
    }
}

/* exported showNewProjectModal */
function showNewProjectModal(hash) {
    loadLocalProjectFromStorage(hash);
    newProjectModal.show();
}

// Slightly modified from: https://stackoverflow.com/a/1349426/2256700
function createRandomString(length) {
    var result           = [];
    var characters       = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    var charactersLength = characters.length;
    for (let i = 0; i < length; i++) {
        result.push(characters.charAt(Math.floor(Math.random() * charactersLength)));
    }
    return result.join('');
}

/* exported getNewLocalProjectHash */
function getNewLocalProjectHash() {
    let id;
    do {
        id = createRandomString(32);
    } while (localStorage.getItem(LOCALSTORAGE_LOCAL_PROJECT_PREFIX + id) !== null);
    return id;
}
