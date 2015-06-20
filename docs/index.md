## Overview
XL Loader (`xlloader`) is a small python module built to import a defined Excel worksheet and convert the contained data into `Ordered Dictionaries` with the keys and values corresponding to a defined header row and the values of the rows.



Example:

| ID       | Product Name        | Modifier |
|----------|:--------------------|----------|
| 1        | Whizbang 5000       | Instant fame and fortune
| 2a       | Recursive Slingshot | +5 annoyance

Produces:

```
import os
import xlloader
from pprint import pprint

xlpath = os.path.join(os.getcwd(),
                        'xlloader',
                        'examples',
                        'product listing.xlsx')


```
