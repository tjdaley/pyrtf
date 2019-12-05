"""
table.py - Create an RTF table.

Copyright (c) 2019 by Thomas J. Daley, J.D. All Rights Reserved.
"""
from collections import namedtuple


class Table(object):
    """
    Encapsulates an RTF table.
    """
    props = [
        'width',
        'borders',
        'alignment',
        'property',
        'header',
        'hfont',
        'dfont',
        'hcolor',
        'dcolor'
    ]

    Column = namedtuple(
        'Column',
        props,
        defaults=(None,) * len(props)
    )

    def __init__(self, columns: list, data, lmargin: int = 0):
        """
        Instance initializer.

        Args:
            columns (list): A list of Column tuples wherein:
                width = width of column in twips OR a float as a percentage
                borders = Any combination of lrtb indicating left, right, top &
                    bottom.
                alignment = One of (l)eft, (r)ight, (c)enter, or (j)ustified
                property = Either the name of a property of a dict or an index
                    into a list.
                header = column header text
                hfont = header font number (index into fonts table)
                dfont = data font number (index into fonts table)
                hcolor = text color for header text (index into color table)
                dcolor = text color for data text (index into color table)

            data list: If a list of lists, then each inner list is a row in the
                table and *property* is an index into the list selecting data
                for that column. If a list of dicts, then each dict is a row in
                the table and *property* is the name of a property to select
                for that column.

            lmargin = Number of twips from left edge of page to begin
        """
        # If someone wanted a single-column table and failed to put the Column
        # specification in a list, fix that here. (Who would DO that?)
        if not isinstance(columns, list):
            self.columns = [columns]
        else:
            self.columns = columns
        self.data = data
        self.lmargin = lmargin

    def __str__(self):
        """
        Produce RTF to represent the table.
        """
        # Specify column widths
        rtf_widths = self.column_widths()
        cells = self.column_rtf_templates()
        rows = []

        # Format the column headers, if present
        if self.has_headers():
            rtf = self.headers(cells)
            rows.append(
                self.begin_row() +
                rtf_widths +
                rtf +
                self.end_row())

        # Format each row of data.
        for r_idx, row in enumerate(self.data):
            rtf = self.data_row(cells, row)
            rows.append(
                self.begin_row() +
                rtf_widths +
                rtf +
                self.end_row()
            )

        # Put it all together and send back to caller
        return '\n' + '\n'.join(rows) + '\n'

    def begin_row(self):
        """
        Produce RTF to begin a row.
        """
        return '{\\trowd\\trgaph180\n'

    def end_row(self):
        """
        Produce RTF to end a row.
        """
        return '\\row}\n'

    def has_headers(self) -> bool:
        """
        Determine whether this table has a header row.

        Args:
            None.
        Returns:
            (bool): True if at least one column has a header, otherwise False.
        """
        for column in self.columns:
            if column.header is not None:
                return True
        return False

    def headers(self, cells) -> str:
        """
        Produce RTF for column headers.
        """
        headers = []
        for c, cell in enumerate(cells):
            column = self.columns[c]
            if column.header is not None:
                col_rtf = ''
                if column.hfont is not None:
                    col_rtf += '\\f{}'.format(column.hfont)
                if column.hcolor is not None:
                    col_rtg += '\\cf{}'.format(column.hcolor)
                col_rtf += ' ' + column.header
                headers.append(cell % col_rtf)
        return '{' + ''.join(headers) + '}\n'

    def data_row(self, cells, data) -> str:
        """
        Produce RTF for a row of data.

        Args:
            cells (list): List of RTF templates for each data cell.
            data (): Dict or List of data to insert into this row.
        """
        cells_rtf = []
        for c, cell in enumerate(cells):
            column = self.columns[c]
            cell_rtf = ''
            if column.dfont is not None:
                cell_rtf += '\\f{}'.format(column.dfont)
            if column.dcolor is not None:
                cell_rtf += '\\cf{}'.format(column.dcolor)
            cell_rtf += self.data_value(column, data)
            cells_rtf.append(cell % cell_rtf)
        return '{' + ''.join(cells_rtf) + '}\n'

    def data_value(self, column: Column, data) -> str:
        """
        Extract a cell of data from the data store.

        Args:
            column (Column): Specification for this column
            data (): Data for this row

        Returns:
            (str): Data value to insert in this column.
        """
        if isinstance(data, list):
            return data[int(column.property)]
        elif isinstance(data, dict):
            return data[column.property]
        return "#ERR#"

    def column_rtf_templates(self) -> list:
        cols = []
        for column in self.columns:
            col = '{\\pard\\q%s\\intbl' % (column.alignment or 'l')
            for border in column.borders or '':
                if border in 'lrtb':
                    col += '\\brdr%s\\brdrs\\brdrw10\\brsp20' % border
            col += ' %s\\cell}\n'
            cols.append(col)

        return cols

    def column_widths(self) -> str:
        widths = []
        doc_width = 1440 * 6.5  # Assuming 8.5 x 11, portrait paper, 1" margins
        use_percent = False
        use_units = False

        for column in self.columns:
            # Convert a percentage string, e.g. '20%' into a float: 0.2
            if isinstance(column.width, str) and column.width[-1] == '%':
                w = float('.' + column.width[:-1])
            # Convert a str to float
            elif isinstance(column.width, str):
                w = float(column.width)
            # Otherwise, just use the number provided.
            elif isinstance(column.width, (int, float)):
                w = column.width
            else:
                raise ValueError("Column.width must be a percent string, e.g. '20%', an int, or a float")  # NOQA

            # Assume that a value less than one is a decimal (percent)
            if w > 1:
                use_units = True
                w = int(column.width)
            else:
                use_percent = True
                w = w * doc_width

            # Add to our widths list.
            widths.append(w)

        # All have to be percent or all have to be in units (twips)
        if use_percent and use_units:
            raise ValueError("All Column.widths must be the same type, either percent or units (twips)")  # NOQA

        # Coerce widths to fix total width of the page.
        total_width = sum(widths)
        coerced_widths = [w / total_width * (doc_width - self.lmargin) for w in widths]  # NOQA

        # Convert decimals to widths in twips
        if use_percent:
            widths = [w * (doc_width - self.lmargin) for w in coerced_widths]
            coerced_widths = widths

        # Convert widths to extents
        total_width = 0
        extents = []
        for w in coerced_widths:
            extents.append(w + total_width)
            total_width += w

        # Finally, create the column extents specification
        widths = ['\\cellx{}'.format(int(w)) for w in extents]
        return ''.join(widths) + '\n'
