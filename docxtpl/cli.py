#!/usr/bin/env python

import click
import os
from click import ClickException
from docxtpl import DocxTemplate
from docxtpl.parsers import PARSERS

def parse_variable_file(file):
    try:
        file_extension = os.path.splitext(file.name)[1]
        return PARSERS[file_extension](file)
    except AttributeError:
        return dict()
    except KeyError:
        error = "Unkown variable file extension '{}'"
        raise ClickException(error.format(file_extension))

@click.command(context_settings=dict(
    help_option_names=["-h", "--help"],
    ignore_unknown_options=True,
))
@click.argument("template")
@click.option("--verbose", "-V", is_flag=True)
@click.option("--variables", "-v", type=click.File("rb"), help="Read template variables from FILENAME. Built-in parsers are JSON, YAML, TOML and XML.")
@click.option("--output", "-o", help="output file")
def cli(template,variables,output,v):
    doc = DocxTemplate(template)
    context = dict()
    for f in variables:
        context.update(parse_variable_file(f))
    doc.render(context)
    doc.save(output)
