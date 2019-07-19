# Extracting data from Excel file

- Run `xlsxreader.py`
- This script creates python files containting dicts with the data for the "upper level" and the "simutool" from the excelfile in the `v3.2` directory. I
- These files are saved to `../domain_model_1_0/`. This directory must exist, or should be created by the user before running the script.
- CAUTION: As changes should be made in these files directly once the project is live, the `xlsxreader.py` has hardocded paths for the imported excel file and the output files. If the output files need to be stored in a different directory, the path needs to change in the code or the files need to be copied. 