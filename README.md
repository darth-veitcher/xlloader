# xlloader
Excel Sheet to Python Dict converter using openpyxl

XL Loader (`xlloader`) is a small python module built to import a defined Excel worksheet and convert the contained data into `Ordered Dictionaries` with the keys and values corresponding to a defined header row and the values of the rows.



Example:

| ID       | Product Name        | Modifier |
|----------|:--------------------|----------|
| 1        | Whizbang 5000       | Instant fame and fortune
| 2a       | Recursive Slingshot | +5 annoyance

Produces:

```
>>> import os
>>> import xlloader
>>> from pprint import pprint

>>> xlpath = os.path.join(os.getcwd(),
                    'xlloader',
                    'examples',
                    'product listing.xlsx')

# Keep original column order
>>> opprint(xlloader.sheet_to_dict(xlpath))

[OrderedDict([(u'ID', 1), (u'Product_Name', u'Whizbang 5000'), (u'Modifier', u'Instant fame and fortune')]),
 OrderedDict([(u'ID', u'2a'), (u'Product_Name', u'Recursive Slingshot'), (u'Modifier', u'+5 annoyance')])]

# Return a standard dict
>>> pprint(xlloader.sheet_to_dict(xlpath, keep_order=False))

[{u'ID': 1,
  u'Modifier': u'Instant fame and fortune',
  u'Product_Name': u'Whizbang 5000'},
 {u'ID': u'2a',
  u'Modifier': u'+5 annoyance',
  u'Product_Name': u'Recursive Slingshot'}]

```
