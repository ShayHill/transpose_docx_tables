"""Extract text from a docx file, converting tables to contiguous lines of
`header: data` pairs.

SEE LAST LINE OF THE FILE FOR USAGE.

Since this code relies on patching implementation details, it is only guaranteed to
work with docx2python 2.10.1

:author: Shay Hill
:created: 2024-04-12
"""


import docx2python as d2p
from docx2python.iterators import iter_at_depth

# Add this marker as a new row in the beginning of each table. The string is
# meaningless, just something with a near-zero probability of appearing in an actual
# file.
OPEN_TABLE_MARKER = "4e62660e-8342-4222-8530-d4ff65856687"

# Just something to put in the text above and below each table row to make them
# easier to see.
TABLE_BORDER = "-" * 20

# ===============================================================================
# Monkey patch the TagRunner class to add the OPEN_TABLE_MARKER to the beginning of
# each table. Each table will start with a 1-cell row with the OPEN_TABLE_MARKER.
#
# [ # table
#   [ # marker row
#     [ # marker cell
#       OPEN_TABLE_MARKER
#     ]
#  ]
#  # now the normal rows
#  [ [header1] [header2] ] [ [data1] [data2] ]
# ]
# ===============================================================================


def _open_table(self, tree):
    """Create an empty row with a table marker."""
    self.tables.commence_paragraph()
    self.tables.insert_text_as_new_run(OPEN_TABLE_MARKER)
    self.tables.conclude_paragraph()
    self.tables.set_caret(2)
    return True


d2p.docx_text.TagRunner._open_table = _open_table


# ===============================================================================
#   Define how paragraphs and tables will be printed.
# ===============================================================================


def iter_text_paragraphs(not_a_table):
    """Iter paragraphs from a table-level element that is not a table."""
    return iter_at_depth(not_a_table, 3)


def iter_table_paragraphs(table):
    """Iter each row from a table as `header: data`.

    There are plenty of empty tables in a typical docx format. First check that each
    table has at least

    1. the throw-away table marker row,
    2. a row of headers,
    3. then at least one row of data.
    """
    if len(table) <= 2:
        return iter_text_paragraphs(table)
    headers = ["<br>".join(x) for x in table[1]]
    for row in table[2:]:
        row_cells = ["<br>".join(x) for x in row]
        yield "\n".join(
            [TABLE_BORDER]
            + [f"{h}: {c}" for h, c in zip(headers, row_cells)]
            + [TABLE_BORDER]
        )


def iter_text_and_table_paragraphs(docx_tables):
    """Iter paragraphs from a docx export."""
    for table in docx_tables:
        try:
            is_table = table[0][0][0] == OPEN_TABLE_MARKER
        except IndexError:
            is_table = False

        if is_table:
            yield from iter_table_paragraphs(table)
        else:
            yield from iter_text_paragraphs(table)


if __name__ == "__main__":
    with d2p.docx2python(
        "input/ps-33471-000ps001_21.docx", duplicate_merged_cells=True
    ) as docx_content:
        docx_tables = docx_content.body

    extracted_paragraphs = list(iter_text_and_table_paragraphs(docx_tables))

    # uncomment below to write the extracted text to a file
    with open("temp.txt", "w", encoding="utf-8") as f:
        _ = f.write("\n\n".join(extracted_paragraphs))
