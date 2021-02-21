/* eslint-env browser */
/* global bootstrap, config */

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
        let tagdiv = document.getElementById('libraryTagFilterDiv');
        let nodes = document.createDocumentFragment();
        config.projecttags.map(function (tag) {
            let node = document.createElement('div');
            node.className = 'form-check d-inline-block me-3 d-lg-block';
            node.innerHTML = '<input class="form-check-input" type="checkbox" value="' + tag + '" id="addProjectModalTagCheck' + tag + '"><label class="form-check-label" for="addProjectModalTagCheck' + tag + '"> ' + tag + '</label>';
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
    project.tags = new Set();
    for (const match of project.title.matchAll(tagRegex)) {
        project.tags.add(match[1]);
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
        aNode.className = 'list-group-item list-group-item-action ' + Array.from(project.tags.keys(), function (tag) {
            return 'tag_class_' +  tag;
        }).join(' ');
        titleNode.innerText = project.title;
        authorNode.innerText = project.author ? project.author.name : '';
        tagsNode.innerHTML = Array.from(project.tags.keys(), function (tag) {
            return '<span class="badge bg-info text-dark me-1">' + tag + '</span>';
        }).join('');
        nodes.appendChild(document.importNode(template.content, true));
        project.node = nodes.lastElementChild;
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
var newProjectModalElement = document.getElementById('newProjectModal');
var newProjectModal = new bootstrap.Modal(newProjectModalElement, {'backdrop': 'static', 'keyboard': false});

new bootstrap.Popover(projectModalElement.querySelector('.projectauthor'));
new bootstrap.Popover(projectModalElement.querySelector('.projectcommitter'));

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

document.getElementById('addProjectToBomButton').addEventListener('click', function () {
    if (currentlyLoadedProject) {
        /* global addProjectToBom */
        addProjectToBom(currentlyLoadedProject, function () {
            projectModal.hide();
        });
    }
});

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
        newProjectModal.hide();
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
    } else if (window.location.hash == '#newProject') {
        projectModal.hide();
        newProjectModal.show();
    } else {
        projectModal.hide();
        newProjectModal.hide();
        currentlyLoadedProject = null;
        switchTab(window.location.hash);
    }
}
window.addEventListener("hashchange", loadProjectFromHash, false);
loadProjectFromHash();
function restoreHashAfterModalClose() {
    if (document.getElementById('intro-tab-pane').classList.contains('active')) {
        window.location.hash = '#intro';
    } else if (document.getElementById('bib-tab-pane').classList.contains('active')) {
        window.location.hash = '#bib';
    } else if (document.getElementById('bom-tab-pane').classList.contains('active')) {
        window.location.hash = '#boms';
    } else {
        window.location.hash = '';
    }
}
projectModalElement.addEventListener('hidden.bs.modal', restoreHashAfterModalClose);
newProjectModalElement.addEventListener('hidden.bs.modal', restoreHashAfterModalClose);

// From https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Regular_Expressions#escaping
function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
}

function refilterLibraryListGroup() {
    let tagDiv = document.getElementById('libraryTagFilterDiv');

    let projects = library;


    let searchtext = document.getElementById('searchProjectInput').value;
    if (searchtext) {
        let regex = '^' + searchtext.split(/\s+/).map(function (word) {
            return '(?=.*?' + escapeRegExp(word) + ')';
        }).join('');
        regex = new RegExp(regex, 'i');
        projects = projects.filter(function (project) {
            return project.title.match(regex);
        });
    }

    let checked = tagDiv.querySelectorAll('input[type="checkbox"]:checked');
    if (checked.length) {
        let selectedTags = Array.prototype.slice.call(checked).map(function (c) {
            return c.value;
        });
        projects = projects.filter(function (project) {
            for (let tag in selectedTags) {
                if (project.tags.has(selectedTags[tag])) {
                    return true;
                }
            }
            return false;
        });
    }

    let nodes = projects.map(function (project) {
        return project.node;
    });
    let fragment = document.createDocumentFragment();
    for (let node in nodes) {
        fragment.appendChild(nodes[node]);
    }
    let listGroup = document.getElementById('projectListGroup');
    listGroup.replaceChildren(fragment);
}

document.getElementById('libraryTagFilterDiv').addEventListener('change', refilterLibraryListGroup);

var searchInputFilterTimeout = null;
document.getElementById('searchProjectInput').addEventListener('input', function () {
    if (searchInputFilterTimeout) {
        clearTimeout(searchInputFilterTimeout);
    }
    searchInputFilterTimeout = setTimeout(function () {
        searchInputFilterTimeout = null;
        refilterLibraryListGroup();
    }, 200);
});
