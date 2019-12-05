"""
pyrtf.py - Classes for creating simple RTF documents.

Copyright (c) 2019 by Thomas J. Daley, J.D. All Rights Reserved.
"""
from collections import namedtuple
from datetime import datetime
import textwrap

from table import Table

Color = namedtuple('Color', ['red', 'green', 'blue'])


class Prolog(object):
    def __str__(self):
        return '\\rtf1\\ansi\\deff0\n'


class FontTable(object):
    def __init__(self, font_names: list = ['Times New Roman', 'Calibri']):
        self.fonts = font_names

    def __str__(self):
        font_table = ['{{\\f{} {};}}'.format(i, font_name)
                      for i, font_name in enumerate(self.fonts)]
        return '{\\fonttbl ' + ''.join(font_table) + '}\n'


class ColorTable(object):
    def __init__(self):
        self.colors = []

    def add_color(self, color: tuple):
        if color is not None:
            new_color = Color(color[0], color[1], color[2])
            self.colors.append(new_color)

    def __str__(self):
        color_table = ['\\red{}\\green{}\\blue{};'.format(
            color.red,
            color.green,
            color.blue
        ) for color in self.colors]
        return '{\\colortbl;' + ''.join(color_table) + '}\n'


class Information(object):
    def __init__(
        self,
        title: str = '',
        author: str = 'discovery.jdbot.us',
        company: str = 'JDBOT, LLC'
    ):
        self.title = title
        self.author = author
        self.company = company
        self.create_time = datetime.now()
        self.comment = "Created by the Discovery Bot"

    def __str__(self):
        return (
            '{\\info\n' +
            '{\\title %s}\n' % self.title +
            '{\\author %s}\n' % self.author +
            '{\\company %s}\n' % self.company +
            '{\\creatim\\yr%s\\mo%s\\dy%s\\hr%s\\min%s}\n' %
                (self.create_time.year, self.create_time.month,  # NOQA
                 self.create_time.day,
                 self.create_time.hour, self.create_time.minute) +
            '{\\doccomm %s}\n' % self.comment +
            '}'
        )


class Margins(object):
    def __init__(
        self,
        top: float = 1.0,
        right: float = 1.0,
        bottom: float = 1.0,
        left: float = 1.0
    ):
        self.top = int(top * 1440.0)
        self.right = int(right * 1440.0)
        self.bottom = int(bottom * 1440.0)
        self.left = int(left * 1440.0)

    def __str__(self):
        return '\\margt{}\\margr{}\\margb{}\\margl{}\n'.format(
            self.top, self.right, self.bottom, self.left
        )


class TabStops(object):
    def __init__(self, *args: float):
        self.tab_stops = [int(tab * 1440.0) for tab in args]

    def add_tab_stop(self, tab_stop: float):
        if tab_stop is not None:
            self.tab_stops.append(int(tab_stop * 1440.0))

    def __str__(self):
        tabs = ['\\tx{}'.format(t) for t in self.tab_stops]
        return ''.join(tabs) + '\n'


class OtherPreliminaries(object):
    def __str__(self):
        return (
            '\\deflang1033' +  # U.S. English
            '\\plain' +  # Reset all formatting
            '\\widowctrl' +  # Widow and orphan control
            '\\hyphauto' +  # Enable automatic, language-specific hyphenation
            '\\ftnbj '  # Footnotes are real footnotes, not endnotes
        )


class Footer(object):
    def __init__(
        self,
        case_name: str = "[INSERT CASE NAME]",
        cause_number: str = "[INSERT CAUSE NUMBER]",
        title: str = "[INSERT DOCUMENT TITLE HERE]"
    ):
        self.case_name = case_name
        self.cause_number = cause_number
        self.title = title

    def __str__(self):
        """
        Creates a standardized footer as follows:

            * Resets all formatting (plain)
            * Left-aligned paragraph (ql)
            * 11-point (fs22)
            * Bold (b)
            * Second entry from the font table (f1)
            * Center tab in the middle of the line (tqc and tx4680)
            * Right tab at the right margin (tqr and tx9360)
            * Top border (brdrt and brdrs) that is 10 twips thick (brdrw10)
              and separated from the text by 20 twips (brsp20)
        """
        return (
            '{\\footer\\pard\\plain\\ql\\fs22\\b\\tqc\\tx4680\\tqr\\tx9360' +
            '\\f1\\adjustright' +
            '\\brdrt\\brdrs\\brdrw10\\brsp20 ' +
            self.case_name.upper() +
            '\\tab\\tab PAGE \\chpgn\\line \n' +
            'Cause #' + self.cause_number + '\\line \n' +
            self.title + '\\par}\n'
        )


class NewLine(object):
    def __str__(self):
        return '\\line \n'


class NewPage(object):
    def __str__(self):
        return '\\page \n'


class TextRun(object):
    UNDERLINE_SINGLE = ''
    UNDERLINE_DOUBLE = 'db'

    # For shortcuts in setting text properties
    props = (
        'color',
        'bold',
        'italic',
        'underline',
        'all_caps',
        'small_caps',
        'strike',
        'outline'
    )
    Properties = namedtuple('Properties', props, defaults=(False,) * len(props))  # NOQA

    # For MD-ish syntax to RTF
    Replacement = namedtuple('Replacement', ['old', 'new'])
    replacements = [
        # Bold
        Replacement(old=' __', new=' \\b '),
        Replacement(old='__ ', new='\\b0  '),
        Replacement(old='__, ', new='\\b0 , '),
        Replacement(old='__.', new='\\b0 .'),

        # Italics
        Replacement(old=' _', new=' \\i '),
        Replacement(old='_ ', new='\\i0  '),
        Replacement(old='_, ', new='\\i0 , '),
        Replacement(old='_.', new='\\i0 .'),

        # Small Caps
        Replacement(old='[[', new='\\scaps '),
        Replacement(old=']]', new='\\scaps0 '),

        # New Line
        Replacement(old='\\n', new='\\line \n'),

        # Practitioner Notes
        Replacement(old='[NOTE: ', new='[\\b\\cf2 NOTE\\b0\\cf1 :'),
    ]

    def __init__(self, text: str, props: Properties = None):
        myprops = props or TextRun.Properties()

        self.text = self.md2rtf(text)
        self.color = myprops.color
        self.bold = myprops.bold
        self.italic = myprops.italic
        self.underline = myprops.underline
        self.all_caps = myprops.all_caps
        self.small_caps = myprops.small_caps
        self.strike_through = myprops.strike
        self.outline = myprops.outline

    def md2rtf(self, text: str) -> str:
        """
        Does a "lite* conversion of a limited number of Markdown
        conventions to RTF.
        """
        result = str(text)
        for r in TextRun.replacements:
            result = result.replace(r.old, r.new)
        return result

    def __str__(self):
        pre = ''
        post = ''

        if not isinstance(self.color, bool):
            pre += '\\cf{}'.format(self.color)
            post += '\\cf0'

        if self.bold:
            pre += '\\b'
            post += '\\b0'

        if self.italic:
            pre += '\\i'
            post += '\\i0'

        if isinstance(self.underline, str):
            pre += '\\ul{}'.format(self.underline)
            post += '\\ul0'

        if self.all_caps:
            pre += '\\caps'
            post += '\\caps0'

        if self.small_caps:
            pre += '\\scaps'
            post += '\\scaps0'

        if self.strike_through:
            pre += '\\strike'
            post += '\\strike0'

        if self.outline:
            pre += '\\outl'
            post += '\\outl0'

        # return '{} {}{} '.format(pre, self.text, post)
        if not pre:
            return '{%s}\n' % self.text
        return '{%s %s}\n' % (pre, self.text)


class Paragraph(object):
    ALIGN_LEFT = 'l'
    ALIGN_RIGHT = 'r'
    ALIGN_CENTER = 'c'
    ALIGN_JUSTIFY = 'j'

    def __init__(self, double_space: bool = False, alignment: str = 'j'):
        self.text = []
        self.double_space = False
        self.alignment = alignment
        self.is_header = False
        self.indent_first_line = True

    def set_header(self):
        self.is_header = True
        self.indent_first_line = False

    def add_text(self, text: TextRun):
        self.text.append(text)

    def __str__(self):
        texts = [str(t) for t in self.text]
        spacing = ''
        keep = ''
        indent = ''
        if self.double_space:
            spacing = '\\sl480\\slmult1'
        if self.is_header:
            # Try not to page break between this and the next paragraph.
            keep = '\\keepn'
        if self.indent_first_line and not self.is_header:
            indent = '\\fi720'  # Indent first line by one-half inch
        return (
            '{{\\pard{}\\q{} '.format(spacing, self.alignment) +
            keep + indent +
            ''.join(texts) +
            '\\par}\n'
        )


class CaseStyle(object):
    props = (
        'cause_number',
        'county',
        'court_type',
        'court_number',
        'petitioner_name',
        'respondent_name',
        'is_divorce',
        'child_names',
        'sensitive',
        'doc_title',
    )
    CaseInfo = namedtuple('CaseInfo', props, defaults=(None,) * len(props))

    def __init__(self, caseinfo: CaseInfo):
        self.cause_number = caseinfo.cause_number
        self.county = caseinfo.county
        self.court_type = caseinfo.court_type
        self.court_number = caseinfo.court_number
        self.petitioner_name = caseinfo.petitioner_name
        self.respondent_name = caseinfo.respondent_name
        self.is_divorce = caseinfo.is_divorce
        self.sensitive = caseinfo.sensitive
        self.doc_title = caseinfo.doc_title
        if isinstance(caseinfo.child_names, str):
            self.child_names = [caseinfo.child_names]
        if isinstance(caseinfo.child_names, list):
            self.child_names = caseinfo.child_names

    def __new_str__(self):
        lcol = Table.Column(
            width=4680,
            borders='r',
            alignment='l',
            property=0,
        )
        rcol = Table.Column(
            width=4680,
            alignment='l',
            property=1,
        )
        columns = [lcol, rcol]

        left_content = ""
        bold_caps = TextRun.Properties(bold=True, all_caps=True)
        if self.is_divorce:
            t = TextRun('In the Matter of\\nThe Marriage of\\n\\n', bold_caps)
            left_content += str(t)
            t = TextRun(self.petitioner_name, bold_caps)
            left_content += str(t)
            t = TextRun('\\nand\\n', bold_caps)
            left_content += str(t)
            t = TextRun(self.respondent_name, bold_caps)
            left_content += str(t)
            if self.child_names:
                t = TextRun('\\n\\nand ', bold_caps)
                left_content += str(t)

        if self.child_names:
            t = TextRun('In the Interest of\\n', bold_caps)
            left_content += str(t)
            if len(self.child_names) == 1:
                capacity = ', a child'
            else:
                capacity = ', minor children'
            t = TextRun(', '.join(self.child_names), bold_caps)
            left_content += str(t)
            t = TextRun(capacity, bold_caps)
            left_content += str(t)

        # Right column
        right_content = ""
        t = TextRun('In the %s Court\\n\\n' % self.court_type, bold_caps)
        right_content += str(t)
        t = TextRun('%s Court #%s\\n\\n' % (self.court_type, self.court_number), bold_caps)  # NOQA
        right_content += str(t)
        t = TextRun('%s County, Texas' % self.county, bold_caps)
        right_content += str(t)

        data = [[left_content, right_content]]

        # Build the table
        table = Table(columns, data)
        return '{' + str(table) + '}\n'

    def __str__(self):
        parts = []
        bold_caps = TextRun.Properties(bold=True, all_caps=True)

        # Sensitive information warning
        if self.sensitive:
            paragraph = Paragraph(alignment=Paragraph.ALIGN_LEFT)
            paragraph.set_header()
            text = TextRun("This document contains\\nsensitive data", bold_caps)  # NOQA
            paragraph.add_text(text)
            parts.append(str(paragraph))

        # Cause Number
        paragraph = Paragraph(alignment=Paragraph.ALIGN_CENTER)
        paragraph.set_header()
        text = TextRun('Cause No. ', bold_caps)
        paragraph.add_text(text)
        text = TextRun(
            self.cause_number,
            TextRun.Properties(underline=TextRun.UNDERLINE_SINGLE, bold=True)
        )
        paragraph.add_text(text)
        paragraph.add_text(NewLine())
        parts.append(str(paragraph))

        # Table containing the full case style in this format
        #
        # In the matter of           |  In the District Court
        # The Marriage of            |
        #                            |  District Court #469
        # John Doe                   |
        # and                        |  Collin County, Texas
        # Jane Doe                   |
        #                            |
        # And in the interest of     |
        # child 1, and child 2,      |
        # Children                   |
        #
        # Not every case style has every element that is in the left column
        # above.
        # Column definitions
        lcol = Table.Column(
            width=4680,
            borders='r',
            alignment='l',
            property=0,
        )
        rcol = Table.Column(
            width=4680,
            alignment='l',
            property=1,
        )
        columns = [lcol, rcol]

        # Construct Data
        t = ""
        if self.is_divorce:
            t += 'In the Matter of\\nThe Marriage of\\n\\n'
            t += self.petitioner_name
            t += '\\nand\\n'
            t += self.respondent_name
            if self.child_names:
                t += '\\n\\nand '

        if self.child_names:
            t += "In the Interest of\\n"
            if len(self.child_names) == 1:
                capacity = ", a child"
            else:
                capacity = ", minor children"
            t += ", ".join(self.child_names)
            t += capacity
        left_content = str(TextRun(t, bold_caps))

        # Right column
        t = 'In the %s Court\\n\\n' % self.court_type
        t += '%s Court #%s\\n\\n' % (self.court_type, self.court_number)
        t += '%s County, Texas' % self.county
        right_content = str(TextRun(t, bold_caps))
        data = [[left_content, right_content]]

        # Build the Table
        table = Table(columns, data)
        parts.append(str(table))
        # case_style = [begin_row, column_widths]
        # case_style.append(left_cell % left_content)
        # case_style.append(right_cell % right_content)
        # case_style.append(end_row)
        # parts.append(''.join(case_style))

        # Document Title
        p = Paragraph(alignment=Paragraph.ALIGN_CENTER)
        p.add_text(NewLine())
        p.set_header()
        t = TextRun(self.doc_title, bold_caps)
        p.add_text(t)
        p.add_text(NewLine())
        parts.append(str(p))

        return '{' + ''.join(parts) + '}\n'


class SignatureBlock(object):
    Attorney = namedtuple(
        'Attorney',
        [
            'name',
            'bar_no',
            'firm_name',
            'street',
            'csz',
            'telephone',
            'fax',
            'email',
            'role'
        ]
    )

    def __init__(self, attorney: Attorney):
        self.attorney = attorney

    def __str__(self):
        parts = []
        line_template = '{\\pard\\ql\\li4680\\keepn %s\\par}\n'
        underline_template = '{\\pard\\ql\\li4680\\keepn\\brdrt\\brdrs\\brdrw10\\brsp20 %s\\par}\n'  # NOQA
        blank_line = '{\\pard\\keepn\\par}\n'
        parts.append(line_template % '\\line Respectfully,\\line')
        parts.append(line_template % self.attorney.firm_name)
        parts.append(line_template % self.attorney.street)
        parts.append(line_template % self.attorney.csz)
        parts.append(line_template % ("Tel: " + self.attorney.telephone))
        parts.append(line_template % ("Fax: " + self.attorney.fax))
        parts.append(blank_line)
        parts.append(line_template % ("/s/ " + self.attorney.name))
        parts.append(underline_template % self.attorney.name)
        parts.append(line_template % ("State Bar No. " + self.attorney.bar_no))
        parts.append(line_template % self.attorney.email)
        parts.append(blank_line)
        parts.append(line_template % self.attorney.role)
        return ''.join(parts)


class CertificateOfService(object):
    Recipient = namedtuple('Recipient', ['name', 'role', 'method', 'address'])

    def __init__(self, attorney: str, designation: str):
        self.attorney = attorney
        self.designation = designation
        self.recipients = []

    def add_recipient(self, recipient: Recipient):
        self.recipients.append(recipient)

    def __str__(self):
        parts = []
        p = Paragraph(alignment=Paragraph.ALIGN_CENTER)
        p.set_header()
        p.double_space = True
        t = TextRun(
            "Certificate of Service\n",
            TextRun.Properties(bold=True, all_caps=True)
        )
        p.add_text(t)
        parts.append(str(p))

        p = Paragraph(alignment=Paragraph.ALIGN_LEFT)
        t = TextRun(textwrap.dedent(
            """
            I certify that a true and correct copy of this document was served
             on each party or attorney of record in compliance with the Texas
             Rules of Civil Procedure on [*_____*] as follows:
            """)
        )
        p.add_text(t)
        p.add_text(NewLine())
        parts.append(str(p))

        for recipient in self.recipients:
            p = Paragraph(alignment=Paragraph.ALIGN_LEFT)
            t = TextRun(recipient.name + ", " + recipient.role)
            p.add_text(t)
            parts.append(str(p))

            p = Paragraph(alignment=Paragraph.ALIGN_LEFT)
            t = TextRun("Via {} to {}".format(
                recipient.method,
                recipient.address
            ), TextRun.Properties(italic=True))
            p.add_text(t)
            p.add_text(NewLine())
            parts.append(str(p))

        signature = (
            '{\\pard\\par} \n' +  # Blank line
            '{\\pard\\ql\\li4680 /s/ %s \\par}' % self.attorney +  # NOQA Electronic signature
            '{\\pard\\ql\\li4680\\brdrt\\brdrs\\brdrw10\\brsp20 ' +  # NOQA Border for signature
            self.attorney + '\\line ' + self.designation + '\\par}'
        )
        parts.append(signature)

        certificate = (
            '\\page \n' +
            ''.join(parts)
        )

        return certificate


class Document(object):
    def __init__(self, title: str = None, cause_number: str = None, case_name: str = None):  # NOQA
        self.header = Prolog()
        self.font_table = FontTable()
        self.color_table = ColorTable()
        self.docinfo = Information(title=title)
        self.font_size = 14
        self.paper_dimensions = '\\paperh15840\\paperw12240\n'
        self.magins = Margins()
        self.tabs = TabStops(.5, 1.0, 3.0)
        self.footer = Footer(case_name=case_name, cause_number=cause_number, title=title)  # NOQA
        self.preliminaries = OtherPreliminaries()
        self.content_sections = []
        self.title = title
        self.cause_number = cause_number
        self.case_name = case_name

    def add_content(self, content):
        self.content_sections.append(content)

    def __str__(self):
        sections = [str(s) for s in self.content_sections]
        content = ''.join(sections)
        return (
            '{' +
            str(self.header) +
            str(self.font_table) +
            str(self.color_table) +
            str(self.docinfo) +
            '\\fs{}\n'.format(self.font_size * 2) +
            str(self.paper_dimensions) +
            str(self.magins) +
            str(self.tabs) +
            str(self.footer) +
            str(self.preliminaries) +
            content +
            '}'
        )


def main():
    # This is the information we need from our database
    doc_title = "Responses to Requests for Production"
    cause_number = "469-55555-2019"
    footer_desc = "IMMO Doe and Doe"

    signing_attorney = SignatureBlock.Attorney(
        'Thomas J. Daley',
        '24059643',
        'Power Daley PLLC',
        '825 Watters Creek Blvd Ste 395',
        'Allen, TX 75013',
        '972-985-4448',
        '972-985-4449',
        'admin@powerdaley.com',
        'Attorney for Respondent'
    )

    case_info = CaseStyle.CaseInfo(
        cause_number,
        'Collin',
        'District',
        '469',
        'John Doe',
        'Jane Doe',
        True,
        ['Johnny Doe', 'Julie Joe'],
        False,
        doc_title
    )

    # Begin building the document
    document = Document(doc_title, cause_number, footer_desc)
    document.color_table.add_color((255, 0, 0))

    # Case Style
    case_style = CaseStyle(case_info)
    document.add_content(str(case_style))

    # Document Content
    bold_small = TextRun.Properties(bold=True, small_caps=True)
    italics = TextRun.Properties(italic=True)
    p = Paragraph(alignment=Paragraph.ALIGN_JUSTIFY)
    t = TextRun('Ava Paxton Daley', bold_small)
    p.add_text(t)
    t = TextRun('provides the _accompanying_ __responses__ to Petitioner\'s ')
    p.add_text(t)
    t = TextRun('[[Requests for Production and Inspection]] ', italics)
    p.add_text(t)
    t = TextRun('propounded by Petitioner on November 1, 2019.')
    p.add_text(t)
    document.add_content(p)

    # Signature block
    signature = SignatureBlock(signing_attorney)
    document.add_content(signature)

    # Certificate of Service
    certificate = CertificateOfService(
        attorney=signing_attorney.name,
        designation=signing_attorney.role
    )
    recipient = CertificateOfService.Recipient(
        'Nicholas Nuspl',
        'Attorney for Petitioner',
        'electronic service',
        'nick@nuspl.com'
    )
    certificate.add_recipient(recipient)
    recipient = CertificateOfService.Recipient(
        'Mary Stanley-Renouf',
        'Assistant Attorney General',
        'electronic service',
        'mary@oag.com'
    )
    certificate.add_recipient(recipient)
    document.add_content(certificate)

    print(str(document))


if __name__ == '__main__':
    main()
