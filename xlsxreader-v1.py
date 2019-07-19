# This file should be executed from the folder where it is saved so that all paths are evaluated correctly

import xlrd  # Read .xlsx and .xls files https://xlrd.readthedocs.io/en/latest/api.html#xlrd-sheet
import pprint  # pretty print
import datetime
import getopt
import sys

file = "sample-kb-domainmodel.xlsx"
try:
     file = getopt.getopt(sys.argv[1:], "hv")[1][0]
except Exception as e:
    pass

workbook = xlrd.open_workbook(file)

ws_classes = workbook.sheet_by_index(0)
ws_properties = workbook.sheet_by_index(1)
ws_relations = workbook.sheet_by_index(2)
ws_namespaces = workbook.sheet_by_index(3)

version = "3.2"  # needs to be updated

# Dict storing all items of upper-level part of Domain Model as single dictionaries
class_dict_upper = {}
# Dict storing all items of lower-level part of Domain Model as single dictionaries
class_dict_simutool = {}
property_dict = {}  # Dict storing all property information
relations_dict_upper = {}  # Dict storing relationsships of upper_level
relations_dict_simutool = {}  # Dict storing relationsships of simutool_level
namespace_dict = {}  # Dict storing namespaces


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

        if str(p_row[1].value) == title and str(p_row[7].value) == "yes":
            required.append(str(p_row[0].value) + ":" + str(p_row[2].value))
        elif str(p_row[1].value) == title and (str(p_row[7].value) == "no" or str(p_row[7].value) == ""):
            optional.append(str(p_row[0].value) + ":" + str(p_row[2].value))

    item_dict.update({
        "label": "TBox",
        "version": version,
        "description": str(row[4].value),
        "required_properties": sorted(required),
        "optional_properties": sorted(optional),
        "subclass_of": sorted(subclass_of),
        "uri": "http://example.org/tbox/" + title.lower()
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
        "uri": "http://example.org/tbox/" + title.lower()
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

