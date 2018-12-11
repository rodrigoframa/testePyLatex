from pylatex import Document, Section, Subsection, Command, PageStyle, Head, LineBreak, simple_page_number, MiniPage, \
    StandAloneGraphic, Figure, Center, NewPage, LongTable, MultiColumn, MultiRow, Subsubsection, Itemize, Tabu, LongTabu
from pylatex.utils import italic, NoEscape, bold
from pylatex.package import Package
import pylatex.config as cf
import json, os


class Ppra(Document):
    def __init__(self):
        document_options = ['a4paper', '10pt']
        geometry_options = {"textwidth": "162mm", "textheight": "240mm"}
        super().__init__(geometry_options=geometry_options, document_options=document_options)
        
        self.packages.append(Package('babel', 'brazilian'))
        self.packages.append(Package('helvet'))
        # \renewcommand{\familydefault}{\sfdefault}
        self.packages.append(Command("renewcommand",
                                     arguments=[NoEscape(r"\familydefault"), NoEscape(r"\sfdefault")]))
    
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
    
    def quadro_planejamento(self, planejamento_acoes):
        with self.create(LongTabu("|X[1,l]|X[3,l]|X[6,l]|")) as tabela_planejamento:
            tabela_planejamento.add_hline()
            tabela_planejamento.add_row((MultiColumn(3, align='|c|', data=bold('Cronograma de Ações')),))
            tabela_planejamento.add_hline()
            tabela_planejamento.add_row((bold('Prazo'), bold('Ação'), bold('Prioridade')))
            tabela_planejamento.add_hline()
            tabela_planejamento.end_table_header()
            tabela_planejamento.add_hline()
            tabela_planejamento.add_row((MultiColumn(3, align='|r|', data='Continua na próxima página'),))
            tabela_planejamento.add_hline()
            tabela_planejamento.end_table_footer()
            tabela_planejamento.end_table_last_footer()
            for acao in planejamento_acoes:
                metas = acao['metas']
                if not len(metas):
                    continue
                col_meta = MultiRow(len(metas), data=acao['prazo'].capitalize())
                row = [col_meta, acao['acao'], '1. ' + metas[0]]
                tabela_planejamento.add_row(row)
                for meta in metas[1:]:
                    tabela_planejamento.add_row(('', '', (str(metas.index(meta) + 1) + '. ' + meta)))
                tabela_planejamento.add_hline()
    
    def quadro_epis(self, fichas_exposicao):
        with self.create(LongTabu("|X[l]|X[l]|")) as tabela_epi:
            tabela_epi.add_hline()
            tabela_epi.add_row((MultiColumn(2, align='|c|', data=bold('Quadro Demonstrativo de EPIs')),))
            tabela_epi.add_hline()
            tabela_epi.add_row((bold('Função'), bold('EPIs Indicados')))
            tabela_epi.add_hline()
            tabela_epi.end_table_header()
            tabela_epi.add_row((MultiColumn(2, align='|c|', data='Continua na próxima página'),))
            tabela_epi.add_hline()
            tabela_epi.end_table_footer()
            tabela_epi.end_table_last_footer()
            for ficha in fichas_exposicao:
                epis = ficha['epis']
                if not len(epis):
                    continue
                col_cargo = MultiRow(len(epis), data=ficha['cargo'])
                row = [col_cargo, epis[0]]
                tabela_epi.add_row(row)
                for epi in epis[1:]:
                    tabela_epi.add_row(('', epi))
                tabela_epi.add_hline()
    
    def quadro_funcionarios(self, quadro_funcionarios, total_masc, total_fem, total_fun):
        with self.create(LongTabu("|X[3,c]|X[c]|X[c]|")) as tabela_funcionarios:
            tabela_funcionarios.add_hline()
            tabela_funcionarios.add_row(
                (MultiRow(2, data=bold('Função')), MultiColumn(2, align='c|', data=bold('Número de Funcionários'))))
            tabela_funcionarios.add_hline(2, 3)
            tabela_funcionarios.add_row(('', bold('Masculino'), bold('Feminino')))
            tabela_funcionarios.end_table_header()
            tabela_funcionarios.add_row((MultiColumn(3, align='|r|', data='Continua na próxima página'),))
            tabela_funcionarios.add_hline()
            tabela_funcionarios.end_table_footer()
            tabela_funcionarios.add_hline()
            tabela_funcionarios.add_row(
                (MultiRow(2, data=bold('Total Funcionários')), bold(total_masc), bold(total_fem)))
            tabela_funcionarios.add_hline(2, 3)
            tabela_funcionarios.add_row(('', MultiColumn(2, align='c|', data=bold(total_fun))))
            tabela_funcionarios.add_hline()
            tabela_funcionarios.end_table_last_footer()
            for funcao in quadro_funcionarios:
                row = [funcao['funcao'], funcao['qtdMasculino'], funcao['qtdFeminino']]
                tabela_funcionarios.add_hline()
                tabela_funcionarios.add_row(row)
    
    def quadro_fichas(self, fichas_exposicao):
        for ficha in fichas_exposicao:
            with self.create(LongTabu("|X[4,l]|X[6,l]|")) as table_fichas:
                table_fichas.add_row((MultiColumn(2, align='|r|', data='Continua na próxima página'),))
                table_fichas.add_hline()
                table_fichas.end_table_footer()
                table_fichas.end_table_last_footer()
                
                # imagem_cargo = recuperar_imagem(ficha['imagemCargo']) if ficha['imagemCargo'] else ''
                # cargo = StandAloneGraphic(imagem_cargo,
                #                           image_options="width=\linewidth, height=\linewidth, keepaspectratio")
                imagem_logo = os.path.join(os.path.dirname(__file__), 'logo.png')
                logo = StandAloneGraphic(imagem_logo,
                                         image_options="width=\linewidth, height=\linewidth, keepaspectratio")
                
                table_fichas.add_hline()
                table_fichas.add_row([logo, ''])
                table_fichas.add_hline()
                table_fichas.add_row((bold('Local/Posto de Trabalho'), ficha['local']))
                table_fichas.add_hline()
                table_fichas.add_row((bold('Função'), ficha['cargo']))
                table_fichas.add_hline()
                table_fichas.add_row((bold('Cargo'), ficha['cargo']))
                table_fichas.add_hline()
                table_fichas.add_row((bold('Quantidade de Trabalhadores'), ficha['qtdTrabalhadores']))
                table_fichas.add_hline()
                table_fichas.add_row((bold('Descricao da Atividade'), ficha['descricaoAtv']))
                table_fichas.add_hline()
            if len(ficha['riscosOcupacionais']):
                # with self.create(LongTabu("|X[4,l]|X[6,l]|", to=r'\linewidth')) as table_fichas:
                #     table_fichas.add_row((MultiColumn(2, align='|r|', data='Continua na próxima página'),))
                #     table_fichas.add_hline()
                #     table_fichas.end_table_footer()
                #     table_fichas.end_table_last_footer()
                #     table_fichas.add_hline()
                table_fichas.add_row((MultiColumn(2, align='|c|', data=bold('Riscos Ocupacionais')),))
                table_fichas.add_hline()
                for risOcu in ficha['riscosOcupacionais']:
                    table_fichas.add_row((MultiColumn(2, align='|c|', data=bold(risOcu['risco'])),))
                    table_fichas.add_hline()
                    table_fichas.add_row((bold('Agente'), risOcu['agente']))
                    table_fichas.add_hline()
                    table_fichas.add_row((bold('Fonte Geradora/Trajetória'), risOcu['listaFonteGeradora']))
                    table_fichas.add_hline()
                    table_fichas.add_row((bold('Exposição'), risOcu['caracterizacao']))
                    table_fichas.add_hline()
                    table_fichas.add_row((bold('Avaliação Qualitativa'), risOcu['avaliacaoQualitativa']))
                    table_fichas.add_hline()
                    table_fichas.add_row((bold('Avaliação Quantitativa'), risOcu['avaliacaoQuantitativa']))
                    table_fichas.add_hline()
                    table_fichas.add_row((bold('Limite Tolerância'), risOcu['limiteTolerancia']))
                    table_fichas.add_hline()
                    table_fichas.add_row((bold('Dados de comprometimento a saúde'), risOcu['comprometimentoSaude']))
                    table_fichas.add_hline()
                    table_fichas.add_row((bold('Possíveis Danos a Saúde'), risOcu['listaDanosSaude']))
                    table_fichas.add_hline()
                    table_fichas.add_row((bold('Medidas de Controle'), risOcu['listaMedidaControle']))
                    table_fichas.add_hline()
    
    def quadro_responsavel(self, profissionais, data_emissao):
        with self.create(LongTabu("|X[l]|X[l]|")) as tabela_resp:
            tabela_resp.add_hline()
            tabela_resp.add_row((MultiColumn(2, align='|c|', data=bold('Responsável pela elaboração')),))
            tabela_resp.add_hline()
            tabela_resp.add_row((bold('Nome'), profissionais[0]['nome']))
            tabela_resp.add_hline()
            tabela_resp.add_row((bold('Função'), profissionais[0]['titulo']))
            tabela_resp.add_hline()
            tabela_resp.add_row((bold('Data da elaboração'), data_emissao))
            tabela_resp.add_hline()
            tabela_resp.add_row((bold('Revisão Programada'), data_emissao))
            tabela_resp.add_hline()
    
    def quadro_empresa(self, ppra):
        with self.create(Tabu("|X[4,l]|X[6,l]|")) as tabela_empresa:
            tabela_empresa.add_hline()
            tabela_empresa.add_row(
                (MultiColumn(2, align='|c|', data=bold('COMPROVANTE DE INSCRIÇÃO E SITUAÇÃO CADASTRAL')),))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('CNPJ'), ppra['CNPJFilial']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Data de abertura'), ppra['data']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Nome empresarial'), ppra['nomeEmpresa'].capitalize()))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Título do estabelecimento (Nome Fantasia)'), ppra['nomeFilial']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Código e descrição da atividade econômica principal CNAE'), ppra['CNAE']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Grau de risco pela NR'), ppra['grauRisco']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Número de funcionários:'), ppra['numFunc']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Horário de Trabalho'), ppra['horaTrabalho']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Atividade'), ppra['atividade']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Logradouro'), ppra['logradouro']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Número'), ppra['numero']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Complemento'), ppra['complemento']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Bairro'), ppra['bairro']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('CEP'), ppra['cep']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Município'), ppra['cidade']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('UF'), ppra['uf']))
            tabela_empresa.add_hline()
            tabela_empresa.add_row((bold('Telefone'), ppra['telefone']))
            tabela_empresa.add_hline()
    
    def assinatura_tecnicos(self, profissionais):
        with self.create(Itemize()) as itemize:
            itemize.add_item(bold("Elaborado em,"))
            itemize.append(Command("vfill"))
            itemize.append(Command("centerline", "DataAtual: 11/12/2018"))
            itemize.append(Command("vfill"))
            itemize.add_item(bold("Visto, Revisado e Aprovado em, "))
            itemize.append(Command("vfill"))
            with itemize.create(MiniPage(align='c')):
                itemize.append("Data Ontem: 21/12/2018")
            itemize.append(Command("vfill"))
            with itemize.create(MiniPage(align='c')):
                carimbo = r"".join("-" * 40) + r"\\{}\\{}\\{}".format(profissionais[0]['nome'],
                                                                      profissionais[0]['titulo'],
                                                                      profissionais[0]['numConselho'])
                itemize.append(Command("textsc", NoEscape(carimbo)))


if __name__ == '__main__':
    conteudo = open('ex.json').read()
    
    json_ppra = json.loads(conteudo)
    # Document
    # cf.active = cf.Version1()
    doc = Ppra()
    
    doc.gera_cabecalho(json_ppra)
    doc.gera_capa(json_ppra)
    doc.append(NewPage())
    doc.change_page_style("header")
    
    doc.append(Command('tableofcontents'))
    # Call function to add text
    # doc.fill_document()
    
    # unicode \uF0FC \uF0d8
    # doc.change_page_style("header")
    for secao in json_ppra['titulosPPRA']:
        # gera titulo pai Secao
        doc.append(NewPage())
        with doc.create(Section(secao['titulo'])):
            doc.append(secao['descricao'])
            if secao['atividade'] == 'Perfil da Empresa':
                doc.quadro_empresa(json_ppra)
            if secao['atividade'] == 'Quadro de Funcionarios':
                doc.quadro_funcionarios(json_ppra['quadroFuncionarios'], json_ppra['qtdMasculino'],
                                        json_ppra['qtdFeminino'], json_ppra['numFunc'])
            if secao['atividade'] == 'Quadro de EPIs':
                doc.quadro_epis(json_ppra['fichaExposicao'])
            if secao['atividade'] == 'Responsavel pelo PPRA':
                doc.quadro_responsavel(json_ppra['profissionais'], json_ppra['dataEmissao'])
            if secao['atividade'] == 'Assinatura dos Técnicos':
                doc.assinatura_tecnicos(json_ppra['profissionais'])
            for sub_secao in secao['tituloFilho']:
                # gera titulo filho Subsecao
                with doc.create(Subsection(sub_secao['titulo'])):
                    doc.append(sub_secao['descricao'])
                    if sub_secao['atividade'] == 'Acões':
                        doc.quadro_planejamento(json_ppra['planejamentoAcoes'])
                    if sub_secao['atividade'] == 'Ficha de Risco':
                        doc.quadro_fichas(json_ppra['fichaExposicao'])
                    for sub_sub_secao in sub_secao['tituloFilho']:
                        # gera titulo filho Subsubsecao
                        with doc.create(Subsubsection(sub_sub_secao['titulo'])):
                            doc.append(sub_sub_secao['descricao'])
    
    doc.generate_pdf('ppra', clean_tex=False, compiler='pdflatex')
    tex = doc.dumps()  # The document as string in LaTeX syntax
