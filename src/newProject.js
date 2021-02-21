/* eslint-env browser */
/* global bootstrap, config, jexcel */

/* exported newProjectModalElement, newProjectModal, */
var newProjectModalElement = document.getElementById('newProjectModal');
var newProjectModal = new bootstrap.Modal(newProjectModalElement, {'backdrop': 'static', 'keyboard': false});

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
        minSpareRows: 3,
        columnSorting: false,
        defaultColWidth: 80,
        allowManualInsertColumn: false,
        tableOverflow: true,
        freezeColumns: 5,
        tableHeight: '600px',
        columns: [
            { type: 'dropdown', title: 'Part Category', width: 80, source: [''].concat(config.categorys) },
            { type: 'text', title: 'Value', width: 80 },
            { type: 'text', title: 'Specification', width: 200 },
            { type: 'text', title: 'Note', width: 200 },
            { type: 'text', title: 'Quantity', width: 80 },
        ],
        updateTable: function(instance, cell, c, r, source, value, id) {
            if (c != NEW_PROJECT_CATEGORY_COL && c != NEW_PROJECT_VALUE_COL && c != NEW_PROJECT_SPEC_COL && c != NEW_PROJECT_NOTE_COL && c != NEW_PROJECT_QTY_COL) {
                cell.classList.add('bg-light');
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
}
initializeNewProjectModal();
