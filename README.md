# Transpose tables in docx2python

Example code to transpose tables in docx2python.

```
header_a, header_b, header_c
datum_a1, datum_b1, datum_c1
datum_a2, datum_b2, datum_c2
datum_a3, datum_b3, datum_c3
```

becomes

```
--------------------
header_a: datum_a1
header_b: datum_b1
header_c: datum_c1
--------------------

--------------------
header_a: datum_a2
header_b: datum_b2
header_c: datum_c2
--------------------

--------------------
header_a: datum_a2
header_b: datum_b2
header_c: datum_c2
--------------------
```

See usage and restrictions in `main.py`.
