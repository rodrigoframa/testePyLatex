from pylatex import Document, Section, Subsection, Command, PageStyle, Head, LineBreak, simple_page_number, MiniPage, \
    StandAloneGraphic, Figure, Center, NewPage
from pylatex.utils import italic, NoEscape
from pylatex.package import Package
import json, os


class Ppra(Document):
    def __init__(self, json_ppra):
        document_options = ['a4paper', '12pt']
        geometry_options = {"textwidth": "162mm", "textheight": "240mm"}
        super().__init__(geometry_options=geometry_options, document_options=document_options)
        
        self.packages.append(Package('babel', 'brazilian'))
    
    def gera_cabecalho(self, json_ppra):
        header = PageStyle("header")
        # Create left header
        with header.create(Head("L")) as header_left:
            logo_file = os.path.join(os.path.dirname(__file__), 'logo.png')
            header_left.append(StandAloneGraphic(image_options="width=3cm", filename=logo_file))
        # Create center header
        with header.create(Head("C")) as header_center:
            header_center.append("PPRA 2018")
            header_center.append(LineBreak())
            header_center.append(json_ppra['nomeEmpresa'])
            header_center.append(LineBreak())
            header_center.append(json_ppra['nomeFilial'])
        # Create right header
        with header.create(Head("R")) as header_right:
            header_right.append(simple_page_number())
            header_right.append(LineBreak())
            header_right.append("Novembro/2018")
        
        self.preamble.append(header)
        
    def gera_capa(self, json_ppra):
        self.preamble.append(Command('title', 'PPRA - Programa de Prevenção de Riscos Ambientais'))
        self.preamble.append(Command('author', json_ppra.get('nomeFilial')))
        self.preamble.append(Command('date', ''))

        logo_file = os.path.join(os.path.dirname(__file__), 'logo.png')
        
        with self.create(MiniPage(align='c')):
            self.append(StandAloneGraphic(image_options="width=5cm", filename=logo_file))
            self.append(NoEscape(r'\maketitle'))
            
        self.append(NoEscape(r'\vfill'))
        with self.create(MiniPage(align='c')):
            self.append("São Luís")
            self.append(LineBreak())
            self.append("Novembro/2018")


if __name__ == '__main__':
    conteudo = open('ex.json').read()
    
    json_ppra = json.loads(conteudo)
    # Document
    doc = Ppra(json_ppra)
    
    doc.gera_cabecalho(json_ppra)
    doc.gera_capa(json_ppra)
    doc.append(NewPage())
    doc.change_page_style("header")
    
    # Call function to add text
    # doc.fill_document()
    
    # Add stuff to the document
    with doc.create(Section('A second section')):
        doc.append('Some text.')
    
    doc.generate_pdf('ppra', clean_tex=False)
    tex = doc.dumps()  # The document as string in LaTeX syntax
