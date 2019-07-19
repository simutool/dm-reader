# This file should be executed from the folder where it is saved so that all paths are evaluated correctly

import xlrd  # Read .xlsx and .xls files https://xlrd.readthedocs.io/en/latest/api.html#xlrd-sheet
import pprint  # pretty print
import datetime
import getopt
import sys


base_uri = "http://example.org/tbox/"

def construct_uri(tail_string):
    return base_uri + tail_string.lower()

file = "v3.2/kb-domainmodel-v3.2-workInProgress.xlsx"
try:
     file = getopt.getopt(sys.argv[1:], "hv")[1][0]
except Exception as e:
    pass

workbook = xlrd.open_workbook(file)

ws_classes = workbook.sheet_by_index(0)
ws_properties = workbook.sheet_by_index(1)
ws_relations = workbook.sheet_by_index(2)
ws_namespaces = workbook.sheet_by_index(3)
ws_meta = workbook.sheet_by_index(4)


meta_rows = ws_meta.get_rows()
meta_rows.next()  # omit header row

version = '0.0'
for row in meta_rows:
    version = row[0].value # read version number from domain model

# Dict storing all items of upper-level part of Domain Model as single dictionaries
class_dict_upper = {}
# Dict storing all items of lower-level part of Domain Model as single dictionaries
class_dict_simutool = {}

property_dict_upper = {}  # Dict storing all property information
property_dict_simutool = {}  # Dict storing all property information

relations_dict_upper = {}  # Dict storing relationsships of upper_level
relations_dict_simutool = {}  # Dict storing relationsships of simutool_level
namespace_dict = {}  # Dict storing namespaces


def eval_binary(field):
    field = field.lower()
    if not field or field=="no" or field=="false":
        return False
    else:
        return True

# -------------- Classes --------------

rows = ws_classes.get_rows()  # generator for iteration rows
rows.next()  # omit header row

for row in rows:
    item_dict = {}  # dict for single entries, created newly every iteration
    required = []
    optional = []
    subclass_of = []

    title = str(row[0].value)
    if str(row[1].value) != "NULL":
        subclass_of.append(str(row[1].value))

    if str(row[2].value) == "upper":
        class_dict_upper.update({title: item_dict})
    if str(row[2].value) == "simutool":
        class_dict_simutool.update({title: item_dict})

    # find required and optional properties for each entity in class_dicts
    p_rows = ws_properties.get_rows()
    for p_row in p_rows:

        if str(p_row[1].value) == title and eval_binary(str(p_row[7].value)):
            # required.append(str(p_row[0].value) + ":" + str(p_row[2].value))
            # for readability, we will not store the qualified name here,
            # the namespace can be found in the relations dict
            required.append(str(p_row[2].value))
        elif str(p_row[1].value) == title and not eval_binary(str(p_row[7].value)):
            # optional.append(str(p_row[0].value) + ":" + str(p_row[2].value))
            # for readability, we will not store the qualified name here,
            # the namespace can be found in the relations dict
             optional.append(str(p_row[2].value))

    item_dict.update({
        "label": "TBox",
        "version": version,
        "description": str(row[4].value),
        "required_property": sorted(required),
        "optional_property": sorted(optional),
        "subclass_of": sorted(subclass_of),
        "identifier": construct_uri(title)
    })


# -------------- Relations --------------

rows = ws_relations.get_rows()  # generator for iterating rows
rows.next()  # omit header row

for row in rows:
    item_dict = {}
    title = str(row[1].value)

    if str(row[3].value) == "upper":
        relations_dict_upper.update({title: item_dict})
    if str(row[3].value) == "simutool":
        relations_dict_simutool.update({title: item_dict})

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
    title = str(row[2].value)

    if str(row[3].value) == "upper":
        property_dict_upper.update({title: item_dict})
    if str(row[3].value) == "simutool":
        property_dict_simutool.update({title: item_dict})

    item_dict.update({
        "namespace": str(row[0].value),
        "title": str(row[2].value),
        "xsd_type": str(row[4].value),
        "description": str(row[5].value),
        "unique": str(eval_binary(str(row[8].value))),
        "identifier": construct_uri(title),
        "label": "property",
        "label2": "TBox"
    })

# -------------- Namespaces --------------

rows = ws_namespaces.get_rows()  # generator for iterating rows
rows.next()  # omit header row

for row in rows:
    item_dict = {}

    namespace_dict.update({str(row[0].value): item_dict})

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

