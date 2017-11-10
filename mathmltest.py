from docx import Document
from lxml import etree

string = ('<math xmlns="http://www.w3.org/1998/Math/MathML">'
    '<mi>P</mi>'
    '<mn>2</mn>'
    '<mn>/</mn>'
    '<mn>m</mn>'
    '</math>')

tree = etree.fromstring(string)
xslt = etree.parse('MML2OMML.XSL')

transform = etree.XSLT(xslt)
new_dom = transform(tree)

doc = Document('test.doc')
for line in doc.paragraphs:
    inline = line.runs
    for i in inline:
        if 'SPGRP' in i.text:
            newLine = i.text.replace('SPGRP', '')
            i.text = newLine
            line._element.append(new_dom.getroot())

#paragraph._element.append()
doc.save('outtest.doc')