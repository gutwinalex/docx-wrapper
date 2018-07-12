# A wrapper for the python-docx module

This script will make your life easier when you try to create reports or other documents in the docx format. It uses the python-docx library in the background which needs to be installed prior. 

```
pip install python-docx
```

If you keep the script in the same folder where your script will be located you will need to import it as: 

```
import docx
```

## Example

```

doc = docx.create_document('template.docx', templates\\head.png')
doc = docx.add_title(doc, 'Report')
doc = docx.add_heading_l1(doc, 'Section 1')

doc = docx.add_paragraph_number(doc, 'Test')
doc = docx.add_paragraph_data(doc, 'IP', variable_that_hold_data)

doc, table = docx.add_table_3(doc, 'Medium Grid 1 Accent 1', 'columnA', 'columnB', 'columnC')
for entry in result_dict:
    doc = docx.add_rows_3(doc, table, entry['dataA'], entry['dataB'], entry['dataC'])

doc = docx.pagebreak(doc)

docx.save_docx(doc, 'document.docx')

```