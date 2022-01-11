from docx.oxml import OxmlElement, ns
import os, zipfile, xml.dom.minidom, sys, getopt

class numbering:
    def create_element(self, name):
        return OxmlElement(name)


    def create_attribute(self, element, name, value):
        element.set(ns.qn(name), value)


    def add_page_number(self, run):
        fldChar1 = self.create_element('w:fldChar')
        self.create_attribute(fldChar1, 'w:fldCharType', 'begin')

        instrText = self.create_element('w:instrText')
        self.create_attribute(instrText, 'xml:space', 'preserve')
        instrText.text = "PAGE"

        fldChar2 = self.create_element('w:fldChar')
        self.create_attribute(fldChar2, 'w:fldCharType', 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

    def file_length(self, fileName):
        document = zipfile.ZipFile(fileName)
        dxml = document.read('docProps/app.xml')
        uglyXml = xml.dom.minidom.parseString(dxml)
        page = uglyXml.getElementsByTagName('Pages')[0].childNodes[0].nodeValue

        return page