/* eslint-env browser */
/* global bootstrap, config, jexcel */

/* exported newProjectModalElement, newProjectModal, */
var newProjectModalElement = document.getElementById('newProjectModal');
var newProjectModal = new bootstrap.Modal(newProjectModalElement, {'backdrop': 'static', 'keyboard': false});

/* exported newProjectTable */
var newProjectTable = null;
function initializeNewProjectModal() {
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
}
initializeNewProjectModal();
