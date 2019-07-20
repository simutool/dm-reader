# This file should be executed from the folder where it is saved so that all paths are evaluated correctly

import xlrd  # Read .xlsx and .xls files https://xlrd.readthedocs.io/en/latest/api.html#xlrd-sheet
import pprint  # pretty print
import datetime
import getopt
import sys


base_uri = "http://example.org/tbox/"

def construct_uri(tail_string):
    return base_uri + tail_string.lower()


def get_props(title, required):
    p_rows = ws_properties.get_rows()
    props = []
    for p_row in p_rows:
        if idx(p_row, _prps,_clsz) == title \
           and _bool(idx(p_row, _prps, 'required')) == required:

            props.append(
              idx(
                row=p_row,
                sheet_name=_prps,
                column_name='title')
              .lower())

    return props


def idx(row, sheet_name, column_name):
    index = indx[sheet_name][column_name]
    return str(row[index].value)

# Read all cells in a row and the name of the columns
# and return a dict with {col1_name:cell1_value, col2_name:cell2_value}
def get_payload_dict_of_class_row(row):
    payload_dict = {}
    for title, index in indx[_clsz].items():
        payload_dict[str(title).lower()] = str(row[index].value)
    return payload_dict


file = None #"v3.2/kb-domainmodel-v3.2-workInProgress.xlsx"
try:
     file = getopt.getopt(sys.argv[1:], "hv")[1][0]
except Exception as e:
    pass

workbook = xlrd.open_workbook(file)

_clsz = 'class'
_prps = 'property'
_rls = 'relation'
_nsps = 'namespace'
_meta = 'model-metadata'
 
ws_classes = workbook.sheet_by_name(_clsz)
ws_properties = workbook.sheet_by_name(_prps)
ws_relations = workbook.sheet_by_name(_rls)
ws_namespaces = workbook.sheet_by_name(_nsps)
ws_meta = workbook.sheet_by_name(_meta)

meta_rows = ws_meta.get_rows()
meta_rows.next()  # omit header row

version = '0.0'
for row in meta_rows:
    version = str(row[0].value) # read version number from domain model

# Dict storing all items of upper-level part of Domain Model as single dictionaries
class_dict_upper = []
# Dict storing all items of lower-level part of Domain Model as single dictionaries
class_dict_simutool = []

property_dict_upper = []  # Dict storing all property information
property_dict_simutool = []  # Dict storing all property information

relations_dict_upper = []  # Dict storing relationsships of upper_level
relations_dict_simutool = []  # Dict storing relationsships of simutool_level
namespace_dict = []  # Dict storing namespaces


# ex. {'Name': 0, 'Subclass of': 1, 'Ontology Level': 2}
# a dict to store the column index of each property
# this makes it easy to access any property by name,
# even when the column location of properties changes
indx = {}

header = ws_classes.get_rows().next()
indx[_clsz] = {t.value.lower():header.index(t) for t in header}


header = ws_properties.get_rows().next()
indx[_prps] = {t.value.lower():header.index(t) for t in header}


header = ws_relations.get_rows().next()
indx[_rls] = {t.value.lower():header.index(t) for t in header}


header = ws_namespaces.get_rows().next()
indx[_nsps] = {t.value.lower():header.index(t) for t in header}


def _bool(field):
    field = field.lower()
    if not field or field=="no" or field=="false":
        return False
    else:
        return True

# -------------- Classes --------------

rows = ws_classes.get_rows()  # generator for iteration rows
header =rows.next()  # omit header row

for row in rows:
    item_dict = {}  # dict for single entries, created newly every iteration
    required = []
    optional = []
    subclass_of = []

    # title = str(row[0].value)
    title = idx(row, _clsz, 'title')

    subcls_of = idx(row, _clsz, 'subclass_of')

    if subcls_of != "NULL":
        subclass_of.append(subcls_of)

    ont_lev = idx(row, _clsz, 'ontology_level') 
    if ont_lev == "upper":
        class_dict_upper.append({title: item_dict})
    else:
        class_dict_simutool.append({title: item_dict})
    
    required = get_props(title, required=True)

    optional = get_props(title, required=False)

    item_dict.update({
        "required_property": sorted(required),
        "optional_property": sorted(optional),
        "label": "TBox"

        # "version": version,
        # "description": str(row[4].value),
        # "subclass_of": sorted(subclass_of),
        # "identifier": construct_uri(title)
    })

    item_dict.update(get_payload_dict_of_class_row(row))


# -------------- Relations --------------

rows = ws_relations.get_rows()  # generator for iterating rows
rows.next()  # omit header row

for row in rows:
    item_dict = {}
    title = str(row[1].value).lower()

    if str(row[3].value) == "upper":
        relations_dict_upper.append({title: item_dict})
    if str(row[3].value) == "simutool":
        relations_dict_simutool.append({title: item_dict})

    item_dict.update({
        "from_entity": str(row[0].value),
        "to_entity": str(row[2].value),
        "level": str(row[3].value),
        "namespace": str(row[4].value),
        "description": str(row[5].value),
        "label": "object_property",
        "identifier": construct_uri(title)
    })


# -------------- Properties (new structure, own dict) --------------

rows = ws_properties.get_rows()  # generator for iterating rows
rows.next()  # omit header row

for row in rows:
    item_dict = {}
    title = str(row[2].value).lower()

    if str(row[3].value) == "upper":
        property_dict_upper.append({title: item_dict})
    if str(row[3].value) == "simutool":
        property_dict_simutool.append({title: item_dict})

    item_dict.update({
        "namespace": str(row[0].value),
        "title": title,
        "xsd_type": str(row[4].value),
        "description": str(row[5].value),
        "unique": str(_bool(str(row[8].value))),
        "identifier": construct_uri(title),
        "label": "property",
        "label2": "TBox"
    })

# -------------- Namespaces --------------

rows = ws_namespaces.get_rows()  # generator for iterating rows
rows.next()  # omit header row

for row in rows:
    item_dict = {}

    namespace_dict.append({str(row[0].value): item_dict})

    item_dict.update({
        "uri": str(row[1].value),
        "url": str(row[2].value),
        "comment": str(row[3].value)
    })



# -------------- Save dicts to .py files, add comments at top --------------

with open("upper.py", "w") as out:
    out.write("# The upper domain model is generic across different data management domain applications. \n# Created {} \n".format(
        datetime.datetime.now()))
    if class_dict_upper:
        out.write("\n# Dict storing class info: \nclasses = ")
        pprint.pprint(class_dict_upper, stream=out)
    if relations_dict_upper:
        out.write("\n# Dict storing relation info. \nrelations = ")
        pprint.pprint(relations_dict_upper, stream=out)
    if property_dict_upper:
        out.write("\n# Dict storing property info. \nproperties = ")
        pprint.pprint(property_dict_upper, stream=out)
    if namespace_dict:
        out.write("\n# Dict storing namespace info. \nnamespaces = ")
        pprint.pprint(namespace_dict, stream=out)

with open("simutool.py", "w") as out:
    out.write("  # Low-level details of the domain model. are specific to the target domain, and hence will contain specific information about manufacturing processes, resources, and so on. This information can be changed according to the specifications of the end-users \n# Created {} \n".format(datetime.datetime.now()))
    if class_dict_simutool:
        out.write("\n# Dict storing class info: \nclasses = ")
        pprint.pprint(class_dict_simutool, stream=out)
    if relations_dict_simutool:
        out.write("\n# Dict storing relation info. \nrelations = ")
        pprint.pprint(relations_dict_simutool, stream=out)
    if property_dict_simutool:
        out.write("\n# Dict storing property info. \nproperties = ")
        pprint.pprint(property_dict_simutool, stream=out)



