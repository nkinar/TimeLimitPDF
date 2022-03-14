"""
Copyright © 2022 Nicholas J. Kinar

Permission is hereby granted, free of charge, to any person
obtaining a copy of this software and associated documentation
files (the “Software”), to deal in the Software without
restriction, including without limitation the rights to use,
copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the
Software is furnished to do so, subject to the following
conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.
"""
import os
import pathlib
import dateparser
from PyPDF2 import PdfFileReader, PdfFileWriter
import chevron
import fitz
import click
from tqdm import tqdm

"""
TimeLimitPDF: A simple utility for closing a PDF and hiding the document when time expires.
The security offered by this utility is good for lab-based final exams in a computer lab,
but is not intended to be comprehensive.

Assumptions:
1. The document is used in a controlled environment where Javascript is enabled in the PDF reader.
2. The reader is similar to Adobe Acrobat and the reader supports the Adobe Javascript API.
3. The PDF reader supports OCGs and layers.
4. The source PDF is not encrypted.

Setup Requirements:
pip install PyMuPDF dateparser chevron fitz click tqdm

NOTES:
PyPDF2 requires a local patched version where the cloneDocumentFromReader()
is changed as per the discussion on StackOverflow:
https://stackoverflow.com/questions/55784897/pypdf2-duplicating-pdf-gives-blank-pages
"""

AUTHOR_STR = 'Author: Nicholas J. Kinar <n.kinar@usask.ca>'
ABOUT_STR = 'TimeLimitPDF: A simple utility for closing a PDF and hiding the document when time expires.'
POSTFIX_START = '_'
POSTFIX_DEFAULT = POSTFIX_START + 'timelimited'
ERROR_DEFAULT = 'An error occurred: '
PDF_EXT = '.pdf'
PDF_SEARCH = '*.pdf'
INPUT_STR = 'INPUT: '
OUTPUT_STR = 'OUTPUT: '
DONE_STR = 'DONE.'

PDF_INSERT_TEXT = """
function TurnOnOCGs(doc, page) 
{
    var arr = doc.getOCGs(page);
    if (!arr) return;
    for (var k = 0; k < arr.length; k++)
    {
        arr[k].state = true;
    }
}

function TriggerOCGs()
{
    for(var k = 0; k < this.numPages; k++)
    {
        TurnOnOCGs(this, k);
    }
}

function CheckExpire(verbose, trg)
{
    var current_time = new Date();
    var end_time = new Date({{year}}, {{monthIndex}}, {{day}}, {{hours}}, {{minutes}});
    if(current_time.getTime() > end_time)
    {
        if(verbose) app.alert("The document has expired and cannot be read.");
        this.closeDoc();
    }
    else
    {   
        if(trg) TriggerOCGs();
    }
} 

function CheckExpireNonVerbose()
{
    CheckExpire(false, false);
}

// Call on startup
CheckExpire(true, true);

// Check every 1 s to close the document
timeout = app.setInterval("CheckExpireNonVerbose()", 1000); 
"""


def render_javascript(date):
    d = {
        'year': date.year,
        'monthIndex': date.month-1,  # Javascript months are 0...11
        'day': date.day,
        'hours': date.hour,
        'minutes': date.minute
    }
    txt = chevron.render(PDF_INSERT_TEXT, d)
    return txt
# DONE


def add_javascript(fn_out, fn_in, text):
    pdf_writer = PdfFileWriter()
    pdf_reader = PdfFileReader(open(fn_in, "rb"))

    # for x in range(0, pdf_reader.getNumPages()):
    #     p = pdf_reader.getPage(x)
    #     pdf_writer.addPage(p)
    pdf_writer.cloneDocumentFromReader(pdf_reader)
    pdf_writer.addJS(text)

    with open(fn_out, 'wb') as f:
        pdf_writer.write(f)
# DONE


def add_javascript_time(fn_out, fn_in, date):
    add_javascript(fn_out, fn_in, render_javascript(date))
# DONE


def make_layers(fn_out, fn_in):
    doc_in = fitz.Document(fn_in)
    doc_out = fitz.Document()
    xc = doc_out.add_ocg('hide', on=True)  # hide layers by default
    n = doc_in.page_count
    for k in range(n):
        page_in = doc_in[k]
        page_out = doc_out.new_page()
        page_out.show_pdf_page(page_in.rect, doc_in, pno=k, oc=xc)
    doc_in.close()
    doc_out.save(fn_out)
# DONE


def encode_str_to_bytes(s):
    return s.encode('utf-8')
# DONE


def replace_tag_javascript(fn):
    with open(fn, 'rb') as f:
        data = f.read()
    data = data.replace(encode_str_to_bytes('/OpenAction'), encode_str_to_bytes('/JavaScript'))
    with open(fn, 'wb') as file:
        file.write(data)
# DONE


def make_layers_add_javascript(fn_out, fn_in, time_str):
    date = dateparser.parse(time_str)
    if date is None:
        raise ValueError('The endtime could not be parsed.')
    make_layers(fn_out, fn_in)
    add_javascript_time(fn_out, fn_out, date)
    replace_tag_javascript(fn_out)
# DONE


########################################################################

def get_filename_stem(fn):
    return os.path.splitext(fn)[0]
# DONE


def get_output_filename_default(fn, postfix):
    stem = get_filename_stem(fn) + postfix + PDF_EXT
    return stem
# DONE


def get_current_dir():
    directory = os.getcwd()
    return directory
# DONE


def get_pdf_files_in_dir(dirname):
    files = list(pathlib.Path(dirname).glob(PDF_SEARCH))
    return files
# DONE


def run_over_directory(dirname, endtime, postfix):
    try:
        files = get_pdf_files_in_dir(dirname)
        if not files:
            click.echo('No PDF files were found in the directory.')
            return
        n = len(files)
        for k in tqdm(range(n)):
            outfile = get_output_filename_default(files[k], postfix)
            make_layers_add_javascript(outfile, files[k], endtime)
    except Exception as e:
        click.echo(ERROR_DEFAULT + str(e))
# DONE


@click.command()
@click.option('-i', '--filein', type=str, required=False, help='Input PDF file')
@click.option('-o', '--fileout', type=str, required=False, help='Output PDF file')
@click.option('-d', '--directory', type=str, required=False, help='Directory of PDF to process')
@click.option('-pf', '--postfix', type=str, required=False, help='Postfix to add to output files in directory')
@click.option('-et', '--endtime', type=str, required=False,
              help='Human-readable time string that limits PDF opening and reading')
def run(filein, fileout, directory, postfix, endtime):
    click.echo(ABOUT_STR)
    click.echo(AUTHOR_STR)
    click.echo('Use --help option to obtain help')
    if not endtime:
        click.echo('The endtime must be specified for document expiry')
        return
    if filein and fileout and directory:
        click.echo('The filein and fileout and directory cannot be all specified')
    if not postfix:
        postfix = POSTFIX_DEFAULT
    if not filein and fileout and not directory:
        click.echo('The filein must be specified')

    if fileout:
        if not fileout.endswith(PDF_EXT):
            fileout += PDF_EXT
        if postfix:
            click.echo('Ignoring postfix since output file name has been specified.')

    if filein and not directory:
        if not fileout:
            fileout = get_filename_stem(filein) + postfix + PDF_EXT
        click.echo(INPUT_STR + str(filein))
        click.echo(OUTPUT_STR + str(fileout))
        try:
            make_layers_add_javascript(fileout, filein, endtime)
        except Exception as e:
            click.echo(ERROR_DEFAULT + str(e))

    if not filein and not fileout and not directory:
        click.echo('Processing files in current directory...')
        run_over_directory(get_current_dir(), endtime, postfix)

    if not filein and not fileout and directory:
        click.echo('Processing files in directory: ' + directory)
        run_over_directory(directory, endtime, postfix)

    click.echo(DONE_STR)
# DONE


if __name__ == '__main__':
    run()
