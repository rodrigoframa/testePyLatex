from pylatex import Document, Section, Subsection, Command, PageStyle, Head, LineBreak, simple_page_number
from pylatex.utils import italic, NoEscape
from pylatex.package import Package
import json


class Ppra(Document):
    def __init__(self, json_ppra):
        document_options = ['a4paper', '12pt']
        geometry_options = {"textwidth": "162mm", "textheight": "240mm"}
        super().__init__(geometry_options=geometry_options, document_options=document_options)
        
        self.packages.append(Package('babel', 'brazilian'))
        self.preamble.append(Command('title', 'PPRA \n Programa de Prevenção de Riscos Ambientais'))
        self.preamble.append(Command('author', json_ppra.get('nomeFilial')))
        self.preamble.append(Command('date', ''))
        self.gera_cabecalho(json_ppra)
        self.append(NoEscape(r'\maketitle'))
        
        
    def gera_cabecalho(self, json_ppra):
        header = PageStyle("header")
        # Create left header
        with header.create(Head("L")):
            header.append("Page date: ")
            header.append(LineBreak())
            header.append("R3")
        # Create center header
        with header.create(Head("C")):
            header.append("PPRA 2018")
            header.append(LineBreak())
            header.append(json_ppra['nomeEmpresa'])
            header.append(LineBreak())
            header.append(json_ppra['nomeFilial'])
        # Create right header
        with header.create(Head("R")):
            header.append(simple_page_number())
            header.append(LineBreak())
            header.append("Novembro/2018")
        
        self.preamble.append(header)
        self.change_document_style("header")


if __name__ == '__main__':
    conteudo = open('ex.json').read()

    json_ppra = json.loads(conteudo)
    # Document
    doc = Ppra(json_ppra)
    

    # Call function to add text
    # doc.fill_document()

    # Add stuff to the document
    with doc.create(Section('A second section')):
        doc.append('Some text.')

    doc.generate_pdf('ppra', clean_tex=False)
    tex = doc.dumps()  # The document as string in LaTeX syntax