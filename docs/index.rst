.. python-docx-template documentation master file, created by
   sphinx-quickstart on Thu Mar 12 17:32:17 2015.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to python-docx-template's documentation!
================================================

.. rubric:: Quickstart

To install::

    pip install docxtpl

Usage::

    from docxtpl import DocxTemplate

    doc = DocxTemplate("my_word_template.docx")
    context = { 'company_name' : "World company" }
    doc.render(context)
    doc.save("generated_doc.docx")

Introduction
------------

This package uses 2 major packages :

- python-docx for reading, writing and creating sub documents
- jinja2 for managing tags inserted into the template docx

python-docx-template has been created because python-docx is powerful for creating documents but not for modifying them.

The idea is to begin to create an example of the document you want to generate with microsoft word, it can be as complex as you want :
pictures, index tables, footer, header, variables, anything you can do with word.
Then, as you are still editing the document with microsoft word, you insert jinja2-like tags directly in the document.
You save the document as a .docx file (xml format) : it will be your .docx template file.

Now you can use python-docx-template to generate as many word documents you want from this .docx template and context variables you will associate.

Note : python-docx-template as been tested with MS Word 97, it may not work with other version.

Jinja2-like syntax
------------------

As the Jinja2 package is used, one can use all jinja2 tags and filters inside the word document.
Nevertheless there are some restrictions and extensions to make it work inside a word document:

Restrictions
++++++++++++

The usual jinja2 tags, are only to be used inside a same run of a same paragraph, it can not be used across several paragraphs, table rows, runs.

Note : a 'run' for microsoft word is a sequence of characters with the same style. For example, if you create a paragraph with all characters the same style :
word will create internally only one 'run' in the paragraph. Now, if you put in bold a text in the middle of this paragraph, word will transform the previous 'run' into 3 'runs' (normal - bold - normal).

Extensions
++++++++++

Tags
....

In order to manage paragraphs, table rows, table columns, runs, special syntax has to be used ::

   {%p jinja2_tag %} for paragraphs
   {%tr jinja2_tag %} for table rows
   {%tc jinja2_tag %} for table columns
   {%r jinja2_tag %} for runs

By using these tags, python-docx-template will take care to put the real jinja2 tags at the right place into the document's xml source code.
In addition, these tags also tells python-docx-template to remove the paragraph, table row, table column or run where are located the begin and ending tags and only takes care about what is in between.

Display variables
.................

As part of jinja2, one can used double braces::

   {{ <var> }}

But if ``<var>`` is an RichText object, you must specify that you are changing the actual 'run' ::

   {{r <var> }}

Note the ``r`` right after the openning braces.

**IMPORTANT** : Do not use the ``r`` variable in your template because ``{{r}}`` could be interpreted as a ``{{r``
without variable specified. Nevertheless, in the lastest doxtpl version you can use a bigger variable name starting
with 'r'. For example ``{{render_color}}`` will be interpreted as ``{{ render_color }}`` not as ``{{r ender_color}}``.

Cell color
..........

There is a special case when you want to change the background color of a table cell, you must put the following tag at the very beginning of the cell ::

   {% cellbg <var> %}

`<var>` must contain the color's hexadecimal code *without* the hash sign

Column spanning
...............

If you want to dynamically span a table cell over many column (this is useful when you have a table with a dynamic column count),
you must put the following tag at the very beginning of the cell to span ::

   {% colspan <var> %}

`<var>` must contain an integer for the number of columns to span. See tests/test_files/dynamic_table.py for an example.

Escaping
........

In order to display ``{%``, ``%}``, ``{{`` or ``}}``, one can use ::

   {_%, %_}, {_{ or  }_}

RichText
--------

When you use ``{{ <var> }}`` tag in your template, it will be replaced by the string contained within `var` variable.
BUT it will keep the current style.
If you want to add dynamically changeable style, you have to use both : the ``{{r <var> }}`` tag AND a ``RichText`` object within `var` variable.
You can change color, bold, italic, size and so on, but the best way is to use Microsoft Word to define your own *caracter* style
( Home tab -> modify style -> manage style button -> New style, select ‘Character style’ in the form ), see example in `tests/richtext.py`
Instead of using ``RichText()``, one can use its shortcut : ``R()``
*Important* : When you use ``{{r }}`` it removes the current character styling from your docx template, this means that if
you do not specify a style in ``RichText()``, the style will go back to a microsoft word default style.
This will affect only character styles, not the paragraph styles (MSWord manages this 2 kind of styles).

Inline image
------------

You can dynamically add one or many images into your document (tested with JPEG and PNG files).
just add ``{{ <var> }}`` tag in your template where ``<var>`` is an instance of doxtpl.InlineImage ::

   myimage = InlineImage(tpl,'test_files/python_logo.png',width=Mm(20))

You just have to specify the template object, the image file path and optionnally width and/or height.
For height and width you have to use millimeters (Mm), inches (Inches) or points(Pt) class.
Please see tests/inline_image.py for an example.

Sub-documents
-------------

A template variable can contain a complex and built from scratch with python-docx word document.
To do so, get first a sub-document object from template object and use it as a python-docx document object, see example in `tests/subdoc.py`.

Escaping, newline, new paragraph, Listing
-----------------------------------------

When you use a ``{{ <var> }}``, you are modifying an **XML** word document, this means you cannot use all chars,
especially ``<``, ``>`` and ``&``. In order to use them, you must escape them. There are 3 ways :

   *  ``context = { 'var':R('my text') }`` and ``{{r <var> }}`` in the template (note the ``r``),
   *  ``context = { 'var':'my text'}`` and ``{{ <var>|e }}`` in your word template
   *  ``context = { 'var':escape('my text')}`` and ``{{ <var> }}`` in the template.

The ``RichText()`` or ``R()`` offers newline and new paragraph feature : just use ``\n`` or ``\a`` in the
text, they will be converted accordingly.

See tests/escape.py example for more informations.

Another solution, if you want to include a listing into your document, that is to escape the text and manage \n and \a,
you can use the ``Listing`` class :

in your python code ::

   context = { 'mylisting':Listing('the listing\nwith\nsome\nlines \a and some paragraph \a and special chars : <>&') }

in your docx template just use ``{{ mylisting }}``
With ``Listing()``, you will keep the current character styling (except after a ``\a`` as you start a new paragraph).

Replace docx pictures
---------------------

It is not possible to dynamically add images in header/footer, but you can change them.
The idea is to put a dummy picture in your template, render the template as usual, then replace the dummy picture with another one.
You can do that for all medias at the same time.
Note: the aspect ratio will be the same as the replaced image
Note2 : Specify the filename that has been used to insert the image in the docx template (only its basename, not the full path)

Syntax to replace dummy_header_pic.jpg::

   tpl.replace_pic('dummy_header_pic.jpg','header_pic_i_want.jpg')


The replacement occurs in headers, footers and the whole document's body.


Replace docx medias
-------------------

It is not possible to dynamically add other medias than images in header/footer, but you can change them.
The idea is to put a dummy media in your template, render the template as usual, then replace the dummy media with another one.
You can do that for all medias at the same time.
Note: for images, the aspect ratio will be the same as the replaced image
Note2 : it is important to have the source media files as they are required to calculate their CRC to find them in the docx.
(dummy file name is not important)

Syntax to replace dummy_header_pic.jpg::

   tpl.replace_media('dummy_header_pic.jpg','header_pic_i_want.jpg')


WARNING : unlike replace_pic() method, dummy_header_pic.jpg MUST exist in the template directory when rendering and saving the generated docx. It must be the same
file as the one inserted manually in the docx template.
The replacement occurs in headers, footers and the whole document's body.

Replace embedded objects
------------------------

It works like medias replacement, except it is for embedded objects like embedded docx.

Syntax to replace embedded_dummy.docx::

   tpl.replace_embedded('embdded_dummy.docx','embdded_docx_i_want.docx')


WARNING : unlike replace_pic() method, embdded_dummy.docx MUST exist in the template directory when rendering and saving the generated docx. It must be the same
file as the one inserted manually in the docx template.
The replacement occurs in headers, footers and the whole document's body.

Tables
------

You can span table cells in two ways, horizontally (see tests/dynamic_table.py) by using::

   {% colspan <number of column to span> %}

or vertically within a for loop (see tests/vertical_merge.py)::

   {% vm %}

Jinja custom filters
--------------------

``render()`` accepts ``jinja_env`` optionnal argument : you may pass a jinja environment object.
By this way you will be able to add some custom jinja filters::

    from docxtpl import DocxTemplate
    import jinja2

    doc = DocxTemplate("my_word_template.docx")
    context = { 'company_name' : "World company" }
    jinja_env = jinja2.Environment()
    jinja_env.filters['myfilter'] = myfilterfunc
    doc.render(context,jinja_env)
    doc.save("generated_doc.docx")

Examples
--------

The best way to see how it works is to read examples, they are located in `tests/` directory. Templates and generated .docx files are in `tests/test_files/`.

Share
-----

If you like this project, please rate and share it here : http://rate.re/github/elapouya/python-docx-template


.. rubric:: Functions index

.. currentmodule:: docxtpl

.. rubric:: Functions documentation

.. automodule:: docxtpl
   :members:


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

