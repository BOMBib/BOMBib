import json, os, sys
import jsonschema
import contextlib

if not (2 <= len(sys.argv) <= 3):
    print("Need to supply 2 Arguments. Folder to search and path to output file")
    print("{} [Project Folder] [Library out File]".format(sys.argv[0]))
    sys.exit(1)

if not os.path.isdir(sys.argv[1]) or not os.access(sys.argv[1], os.R_OK):
    print("First argument is not a readable directory")
    sys.exit(2)
directory = sys.argv[1]

if len(sys.argv) > 2:
    if os.path.exists(sys.argv[2]) and not os.access(sys.argv[2], os.W_OK):
        print("Second argument is not writable")
        sys.exit(3)
    outfile = sys.argv[2]
    relativeFromPath = os.path.dirname(outfile)
else:
    outfile = '-'
    relativeFromPath = os.curdir

#From https://stackoverflow.com/a/17603000/2256700
@contextlib.contextmanager
def smart_open_write(filename=None):
    if filename and filename != '-':
        fh = open(filename, 'w')
    else:
        fh = sys.stdout

    try:
        yield fh
    finally:
        if fh is not sys.stdout:
            fh.close()


PROJECTFILENAME = 'bomproject.json'

person_schema = {
    "type": "object",
    "properties": {
        "name": { "type": "string" },
        "github": { "type": "string" }, #TODO: Validate URL
        "patreon": { "type": "string" }, #TODO: Validate URL
        "youtube": { "type": "string" }, #TODO: Validate URL
        "web": { "type": "string" }
    },
    "required": ["name"]
}

schema = {
        "type": "object",
        "properties": {
            "title": { "type": "string" },
            "description": { "type": "string" },
            "author": person_schema,
            "committer": person_schema,
            "bom": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "type" : { "type": "string" }, #TODO: Validate type with list of allowed types
                        "value" : { "type": "string" },
                        "spec" : { "type": "string" },
                        "qty" : { "type": "integer" },
                        "note": { "type": "string" }
                    },
                    "required": ["spec", "qty"]
                }
            }
        },
        "required": ["title", "author", "committer"]
}
validator = jsonschema.validators.Draft3Validator(schema)
with smart_open_write(outfile) as outhandle:
    library = []
    for root, subs, files in os.walk(directory):
        if PROJECTFILENAME in files:
            path = os.path.join(root, PROJECTFILENAME)
            with open(path, "r") as h:
                project = json.load(h)
                validator.validate(instance=project)
                library.append({
                    "t": project["title"],
                    "a": project["author"]["name"],
                    "p": os.path.relpath(path, relativeFromPath)
                })
    json.dump(library, outhandle)
