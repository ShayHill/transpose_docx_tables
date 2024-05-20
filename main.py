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

# When header rows are combined, separate them with this string. Using an underscore
# padded with space to it won't be interpreted as markdown.
HEADER_SEPARATOR = " _ "

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


def _are_all_unique(seq):
    """Check if all elements in a sequence are unique."""
    return len(seq) == len(set(seq))


def _join_table_cell(cell):
    """Join a cell into a single string."""
    return "<br>".join(cell)


def _join_table_row_cells(row):
    """Join each cell in a row into a single string."""
    return [_join_table_cell(cell) for cell in row]


def _combine_headers(table):
    """Combine headers into a single string.

    Many tables are built

    | header1  | header2  | header3  |
    | ---------|----------|----------|
    | datum1   | datum2   | datum3   |

    Other tables are built

    | CATEGORY | CATEGORY | CATEGORY |
    | ---------|----------|----------|
    | header1  | header2  | header3  |
    | datum1   | datum2   | datum3   |

    Keep combining header rows until each column has a unique header. E.g.,
    "CATEGORY _ header1", "CATEGORY _ header2", "CATEGORY _ header3"

    Return the combined headers. This function has a side effect of removing the
    header rows from the input table.
    """
    headers = _join_table_row_cells(table.pop(0))
    while not _are_all_unique(headers):
        try:
            next_header_row = _join_table_row_cells(table.pop(0))
        except IndexError:
            msg = "Cannot create a unique header for each column."
            raise ValueError(msg)
        headers = [HEADER_SEPARATOR.join(h) for h in zip(headers, next_header_row)]
    return headers


def iter_table_paragraphs(table):
    """Iter each row from a table as `header: data`.

    There are plenty of empty tables in a typical docx format. First check that each
    table has at least

    1. the throw-away table marker row,
    2. a row of headers,
    3. then at least one row of data.
    """
    # strip first row, which is the table marker
    table = table[1:]

    # a table with headers only and no data is most likely just something the docx
    # author used to format the document. Print the contents as text.
    if len(table) <= 1:
        return iter_text_paragraphs(table)

    headers = _combine_headers(table)
    for row in table:
        row_cells = _join_table_row_cells(row)
        yield "\n".join(
            [TABLE_BORDER]
            + [f"{h}: {c}" for h, c in zip(headers, row_cells)]
            + [TABLE_BORDER]
        )


def iter_text_and_table_paragraphs(docx_tables):
    """Iter paragraphs from a docx export."""
    for potential_table in docx_tables:
        try:
            is_table = potential_table[0][0][0] == OPEN_TABLE_MARKER
        except IndexError:
            is_table = False

        if is_table:
            yield from iter_table_paragraphs(potential_table)
        else:
            yield from iter_text_paragraphs(potential_table)


if __name__ == "__main__":
    with d2p.docx2python(
        "input/ps-33471-000ps001_21.docx", duplicate_merged_cells=True
    ) as docx_content:
        docx_tables = docx_content.body

    extracted_paragraphs = list(iter_text_and_table_paragraphs(docx_tables))

    # uncomment below to write the extracted text to a file
    with open("temp.txt", "w", encoding="utf-8") as f:
        _ = f.write("\n\n".join(extracted_paragraphs))
