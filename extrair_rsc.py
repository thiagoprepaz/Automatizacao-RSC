"""Consolida lançamentos RSC de uma pasta de arquivos ODS.

Python 3.9+, somente biblioteca padrão. Não executa macros nem fórmulas.
"""
import argparse
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import re
import sys
import unicodedata
from xml.dom import minidom
from zipfile import ZipFile
from datetime import datetime
from tempfile import TemporaryDirectory
import threading
import queue
import time

TABLE = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'
OFFICE = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
TEXT = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
BASE = Path('O:/CAPP/DPP - DIVISÃO DE PAGAMENTO DE PESSOAL/_SEPAR/Lançamentos/RSC-Reconhecimento de Saberes e Competências/Projeto de Automatização - Macro RSC')
MAPA = {'SIAPE/7': 'B8', 'Nome': 'B7', 'MÊS/24': 'B10',
        'Rubrica': 'E15', 'R/D': None, 'SEQ': 'B11',
        'Valor/0000': 'I18', 'Justificativa': 'I11', 'Doc legal': 'I10'}


def elementos(no, nomes):
    return [n for n in no.childNodes if n.nodeType == n.ELEMENT_NODE
            and n.namespaceURI == TABLE and n.localName in nomes]


def linhas(no):
    for filho in elementos(no, {'table-row', 'table-header-rows', 'table-row-group', 'table-rows'}):
        if filho.localName == 'table-row':
            yield filho
        else:
            yield from linhas(filho)


def repeticoes(no, atributo):
    return int(no.getAttributeNS(TABLE, atributo) or '1')


def localizar(nos, indice, atributo):
    inicio = 0
    for no in nos:
        tamanho = repeticoes(no, atributo)
        if inicio <= indice < inicio + tamanho:
            return no, indice - inicio, tamanho
        inicio += tamanho
    return None, indice - inicio, 0


def separar(no, deslocamento, tamanho, atributo):
    pai = no.parentNode
    if deslocamento:
        antes = no.cloneNode(True)
        antes.setAttributeNS(TABLE, 'table:' + atributo, str(deslocamento))
        pai.insertBefore(antes, no)
    if tamanho - deslocamento - 1:
        depois = no.cloneNode(True)
        depois.setAttributeNS(TABLE, 'table:' + atributo, str(tamanho - deslocamento - 1))
        pai.insertBefore(depois, no.nextSibling)
    if no.hasAttributeNS(TABLE, atributo):
        no.removeAttributeNS(TABLE, atributo)
    return no


def celula(aba, linha, coluna, editar=False):
    no, deslocamento, tamanho = localizar(linhas(aba), linha, 'number-rows-repeated')
    if no is None:
        if not editar:
            return None
        for _ in range(deslocamento + 1):
            no = aba.ownerDocument.createElementNS(TABLE, 'table:table-row')
            aba.appendChild(no)
    elif editar:
        no = separar(no, deslocamento, tamanho, 'number-rows-repeated')
    c, deslocamento, tamanho = localizar(elementos(no, {'table-cell', 'covered-table-cell'}), coluna, 'number-columns-repeated')
    if c is None:
        if not editar:
            return None
        for _ in range(deslocamento + 1):
            c = aba.ownerDocument.createElementNS(TABLE, 'table:table-cell')
            no.appendChild(c)
    elif editar:
        c = separar(c, deslocamento, tamanho, 'number-columns-repeated')
        if c.localName == 'covered-table-cell':
            raise ValueError('A célula de destino faz parte de uma mesclagem.')
    return c


def texto(no):
    if no.nodeType in (no.TEXT_NODE, no.CDATA_SECTION_NODE):
        return no.data
    if no.namespaceURI == TEXT:
        if no.localName == 's':
            return ' ' * int(no.getAttributeNS(TEXT, 'c') or '1')
        if no.localName == 'tab':
            return '\t'
        if no.localName == 'line-break':
            return '\n'
    return ''.join(texto(n) for n in no.childNodes)


def valor(c):
    if c is None:
        return ''
    tipo = c.getAttributeNS(OFFICE, 'value-type')
    if tipo in ('float', 'currency', 'percentage'):
        bruto = c.getAttributeNS(OFFICE, 'value')
        if bruto:
            numero = Decimal(bruto)
            if not numero.is_finite():
                raise ValueError('Valor numérico inválido na planilha.')
            return numero
    if tipo == 'date':
        return ('date', c.getAttributeNS(OFFICE, 'date-value'))
    ps = c.getElementsByTagNameNS(TEXT, 'p')
    resultado = '\n'.join(texto(p) for p in ps)
    if not resultado:
        resultado = c.getAttributeNS(OFFICE, 'string-value')
    if c.hasAttributeNS(TABLE, 'formula') and not resultado:
        raise ValueError('Fórmula sem resultado salvo. Recalcule e salve a origem no LibreOffice.')
    return resultado


def ler(aba, endereco):
    letras, numero = re.fullmatch(r'([A-Z]+)([1-9][0-9]*)', endereco).groups()
    coluna = 0
    for letra in letras:
        coluna = coluna * 26 + ord(letra) - 64
    return valor(celula(aba, int(numero) - 1, coluna - 1))


def normalizar(v):
    s = unicodedata.normalize('NFKD', str(v)).casefold()
    return ''.join(c for c in s if c.isalnum())


def numero(v, endereco='I18'):
    if isinstance(v, Decimal):
        return v
    s = str(v).strip().replace('R$', '').replace('\u00a0', '').replace(' ', '')
    if not s:
        return Decimal(0)
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        n = Decimal(s)
        if not n.is_finite():
            raise InvalidOperation
        return n
    except InvalidOperation:
        raise ValueError(f'{endereco} não contém um número válido. Recalcule e salve a origem.') from None


def abrir(caminho):
    with ZipFile(caminho) as z:
        doc = minidom.parseString(z.read('content.xml'))
    planilha = doc.getElementsByTagNameNS(OFFICE, 'spreadsheet')[0]
    return doc, elementos(planilha, {'table'})


def escolher(abas, nome):
    if nome is None:
        return abas[0]
    for aba in abas:
        if aba.getAttributeNS(TABLE, 'name') == nome:
            return aba
    raise ValueError('Aba não encontrada: ' + nome)


def cabecalhos(abas):
    esperados = {normalizar(k): k for k in MAPA}
    encontrados = []
    for aba in abas:
        for r in range(100):
            campos = {}
            for c in range(100):
                chave = normalizar(valor(celula(aba, r, c)))
                if chave in esperados:
                    nome = esperados[chave]
                    if nome in campos:
                        raise ValueError('Cabeçalho duplicado: ' + nome)
                    campos[nome] = c
            if len(campos) == len(MAPA):
                encontrados.append((aba, r, campos))
    if len(encontrados) != 1:
        raise ValueError('Esperada uma linha com todos os cabeçalhos nas primeiras 100 linhas/colunas. Encontradas: ' + str(len(encontrados)))
    return encontrados[0]


def escrever(c, v):
    for atributo in list(c.attributes.values()):
        if atributo.namespaceURI == OFFICE or (atributo.namespaceURI == TABLE and atributo.localName == 'formula'):
            c.removeAttributeNode(atributo)
    for filho in list(c.childNodes):
        c.removeChild(filho)
    if isinstance(v, Decimal):
        c.setAttributeNS(OFFICE, 'office:value-type', 'float')
        c.setAttributeNS(OFFICE, 'office:value', str(v))
        s = format(v, 'f')
    elif isinstance(v, tuple) and v[0] == 'date':
        c.setAttributeNS(OFFICE, 'office:value-type', 'date')
        c.setAttributeNS(OFFICE, 'office:date-value', v[1])
        s = v[1]
    else:
        c.setAttributeNS(OFFICE, 'office:value-type', 'string')
        s = str(v)
    for parte in s.split('\n'):
        p = c.ownerDocument.createElementNS(TEXT, 'text:p')
        p.appendChild(c.ownerDocument.createTextNode(parte))
        c.appendChild(p)


def executar(origem, modelo, saida, aba_origem=None, aba_destino=None):
    origem, modelo, saida = map(Path, (origem, modelo, saida))
    if saida.resolve() in (origem.resolve(), modelo.resolve()):
        raise ValueError('A saída precisa ser diferente dos arquivos originais.')
    if saida.exists():
        raise ValueError('A saída já existe. Escolha outro nome para evitar sobrescrever.')
    _, abas = abrir(origem)
    fonte = escolher(abas, aba_origem)
    lancamentos = []
    for endereco, rubrica, tipo in [('I18', 'E15', 'r'), ('I24', 'E21', 'd')]:
        total = numero(ler(fonte, endereco), endereco)
        if total > 0:
            dados = {k: ler(fonte, ref) for k, ref in MAPA.items()
                     if k not in ('Rubrica', 'R/D', 'Valor/0000')}
            dados['Rubrica'] = ler(fonte, rubrica)
            dados['R/D'] = tipo
            dados['Valor/0000'] = total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            lancamentos.append(dados)
    if not lancamentos:
        return 0
    doc, abas = abrir(modelo)
    if aba_destino:
        abas = [escolher(abas, aba_destino)]
    destino, header, campos = cabecalhos(abas)
    # Usa a primeira linha inteiramente vazia após os cabeçalhos.
    linha = header + 1
    for dados in lancamentos:
        while linha < 1048576:
            if all(valor(celula(destino, linha, c)) == '' and
                   not (celula(destino, linha, c) and celula(destino, linha, c).hasAttributeNS(TABLE, 'formula'))
                   for c in campos.values()):
                break
            linha += 1
        else:
            raise ValueError('Não há linha vazia disponível.')
        for nome, c in campos.items():
            escrever(celula(destino, linha, c, editar=True), dados[nome])
        linha += 1
    saida.parent.mkdir(parents=True, exist_ok=True)
    # Mantém estilos, outras abas e demais componentes do modelo ODS.
    with ZipFile(modelo) as entrada, saida.open('xb') as arquivo:
        with ZipFile(arquivo, 'w') as resultado:
            for item in entrada.infolist():
                conteudo = doc.toxml(encoding='utf-8') if item.filename == 'content.xml' else entrada.read(item.filename)
                resultado.writestr(item, conteudo)
    return len(lancamentos)


def criar_estrutura(caminho):
    """Cria a estrutura ODS de saída, sem depender de um modelo externo."""
    from xml.sax.saxutils import escape
    mime = 'application/vnd.oasis.opendocument.spreadsheet'
    cabecalho = ''.join(
        '<table:table-cell office:value-type="string"><text:p>' + escape(nome) +
        '</text:p></table:table-cell>' for nome in MAPA)
    conteudo = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" '
        f'xmlns:text="{TEXT}" office:version="1.2">'
        '<office:body><office:spreadsheet><table:table table:name="Lançamentos">'
        '<table:table-row>' + cabecalho + '</table:table-row>'
        '</table:table></office:spreadsheet></office:body></office:document-content>')
    manifest = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">'
        f'<manifest:file-entry manifest:full-path="/" manifest:media-type="{mime}"/>'
        '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
        '</manifest:manifest>')
    with ZipFile(caminho, 'w') as z:
        z.writestr('mimetype', mime)
        z.writestr('content.xml', conteudo.encode('utf-8'))
        z.writestr('META-INF/manifest.xml', manifest.encode('utf-8'))


def formatar_saida(origem, saida):
    """Ajusta larguras ao conteúdo e exibe sempre exatamente 2 casas decimais."""
    doc, abas = abrir(origem)
    aba = abas[0]
    STYLE = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'
    FO = 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'
    NUMBER = 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0'
    for prefixo, uri in [('style', STYLE), ('fo', FO), ('number', NUMBER)]:
        doc.documentElement.setAttribute('xmlns:' + prefixo, uri)
    estilos = doc.createElementNS(OFFICE, 'office:automatic-styles')
    corpo = doc.getElementsByTagNameNS(OFFICE, 'body')[0]
    doc.documentElement.insertBefore(estilos, corpo)

    def elemento(pai, uri, nome, attrs):
        no = doc.createElementNS(uri, nome)
        for chave, v in attrs.items():
            no.setAttribute(chave, str(v))
        pai.appendChild(no)
        return no

    formato = elemento(estilos, NUMBER, 'number:number-style', {'style:name': 'DuasCasas', 'number:language': 'pt', 'number:country': 'BR'})
    elemento(formato, NUMBER, 'number:number', {'number:decimal-places': 2, 'number:min-decimal-places': 2, 'number:min-integer-digits': 1})
    for nome in ('Texto', 'Valor', 'Cabecalho'):
        attrs = {'style:name': nome, 'style:family': 'table-cell'}
        if nome == 'Valor':
            attrs['style:data-style-name'] = 'DuasCasas'
        estilo = elemento(estilos, STYLE, 'style:style', attrs)
        props = {'fo:wrap-option': 'wrap', 'fo:padding': '0.12cm', 'style:vertical-align': 'middle'}
        if nome == 'Cabecalho':
            props['fo:background-color'] = '#18344d'
        elemento(estilo, STYLE, 'style:table-cell-properties', props)
        elemento(estilo, STYLE, 'style:text-properties', {'fo:font-family': 'Calibri', 'fo:font-size': '11pt', 'fo:color': '#ffffff' if nome == 'Cabecalho' else '#172b3a', 'fo:font-weight': 'bold' if nome == 'Cabecalho' else 'normal'})
    estilo = elemento(estilos, STYLE, 'style:style', {'style:name': 'Linha', 'style:family': 'table-row'})
    elemento(estilo, STYLE, 'style:table-row-properties', {'style:use-optimal-row-height': 'true', 'style:min-row-height': '0.7cm'})
    larguras = [len(k) for k in MAPA]
    for r, linha in enumerate(linhas(aba)):
        linha.setAttributeNS(TABLE, 'table:style-name', 'Linha')
        for c, cel in enumerate(elementos(linha, {'table-cell'})):
            if c >= len(larguras):
                continue
            v = valor(cel)
            exibido = format(v, 'f') if isinstance(v, Decimal) else str(v)
            larguras[c] = max(larguras[c], max(map(len, exibido.splitlines()), default=0))
            cel.setAttributeNS(TABLE, 'table:style-name', 'Cabecalho' if r == 0 else ('Valor' if c == 6 else 'Texto'))
    primeira = aba.firstChild
    for c, tamanho in enumerate(larguras):
        estilo = elemento(estilos, STYLE, 'style:style', {'style:name': f'Coluna{c}', 'style:family': 'table-column'})
        # Textos muito longos quebram em linhas, mantendo a coluna legível.
        elemento(estilo, STYLE, 'style:table-column-properties', {'style:column-width': f'{max(2.5, min(16, (tamanho + 3) * 0.23)):.2f}cm'})
        coluna = doc.createElementNS(TABLE, 'table:table-column')
        coluna.setAttributeNS(TABLE, 'table:style-name', f'Coluna{c}')
        aba.insertBefore(coluna, primeira)
    with ZipFile(origem) as entrada, saida.open('xb') as arquivo:
        with ZipFile(arquivo, 'w') as z:
            for item in entrada.infolist():
                z.writestr(item, doc.toxml(encoding='utf-8') if item.filename == 'content.xml' else entrada.read(item.filename))


def processar_pasta(pasta, progresso=lambda mensagem: None, avancar=lambda percentual: None):
    """Processa todos os ODS diretamente na pasta e consolida os aprovados."""
    inicio = time.monotonic()
    enviar_mensagem = progresso
    ultima_mensagem = inicio - 0.4
    intervalo = 0.4

    def progresso(mensagem):
        nonlocal ultima_mensagem
        espera = intervalo - (time.monotonic() - ultima_mensagem)
        if espera > 0:
            time.sleep(espera)
        enviar_mensagem(mensagem)
        ultima_mensagem = time.monotonic()

    progresso('Iniciando leitura da pasta…')
    avancar(0)
    pasta = Path(pasta)
    if not pasta.is_dir():
        raise ValueError('Selecione uma pasta válida.')
    arquivos = sorted((p for p in pasta.iterdir()
                       if p.is_file() and p.suffix.lower() == '.ods'
                       and p.name.casefold() != 'modelo para entrada macro.ods'
                       and not p.name.startswith('~$')
                       and not p.name.lower().startswith(('entrada macro preenchida', 'entrada para macro'))),
                      key=lambda p: p.name.casefold())
    if not arquivos:
        raise ValueError('Nenhuma planilha ODS de origem encontrada na pasta.')
    intervalo = 8 / (2 * len(arquivos) + 5)
    meses = ('janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro')
    agora = datetime.now()
    resultado_dir = pasta / f'Lançamentos RSC - {meses[agora.month - 1]} {agora.year}'
    resultado_dir.mkdir(parents=True, exist_ok=True)
    progresso(f'{len(arquivos)} planilha(s) encontrada(s).')
    avancar(5)
    nome_relatorio = f'Relatório de Execução - RSC lançamento {agora:%d-%m-%Y}'
    relatorio = resultado_dir / f'{nome_relatorio}.txt'
    versao_relatorio = 2
    while relatorio.exists():
        relatorio = resultado_dir / f'{nome_relatorio} ({versao_relatorio}).txt'
        versao_relatorio += 1
    registros = []
    arquivos_com_erro = []
    planilhas_incluidas = 0
    incluidos = ignorados = erros = 0
    # Cada arquivo é aplicado a uma cópia temporária: uma falha não deixa
    # um lançamento parcialmente preenchido no resultado consolidado.
    with TemporaryDirectory(prefix='rsc_') as temporaria:
        modelo = Path(temporaria) / 'estrutura.ods'
        criar_estrutura(modelo)
        atual = modelo
        for i, origem in enumerate(arquivos, 1):
            progresso(f'Lendo {i}/{len(arquivos)}: {origem.name}')
            proximo = Path(temporaria) / f'lote_{i}.ods'
            try:
                quantidade = executar(origem, atual, proximo)
                if proximo.exists():
                    anterior = atual
                    atual = proximo
                    incluidos += quantidade
                    planilhas_incluidas += 1
                    registros.append(f'INCLUÍDO: {origem.name} — {quantidade} lançamento(s)')
                    if anterior != modelo:
                        anterior.unlink()
                else:
                    ignorados += 1
                    registros.append(f'SEM LANÇAMENTO: {origem.name} — I18 e I24 não possuem valor maior que zero.')
            except Exception as exc:
                erros += 1
                arquivos_com_erro.append(origem.name)
                registros.append(f'ERRO: {origem.name} — {exc}')
            progresso(registros[-1])
            avancar(5 + 80 * i / len(arquivos))
        if incluidos:
            nome_saida = f'Entrada para macro - {planilhas_incluidas} planilhas incluídas'
            saida = resultado_dir / f'{nome_saida}.ods'
            versao = 2
            while saida.exists():
                saida = resultado_dir / f'{nome_saida} ({versao}).ods'
                versao += 1
            progresso('Ajustando colunas e formatando valores…')
            formatar_saida(atual, saida)
            progresso(f'Planilha salva: {saida.name}')
        avancar(95)
    resumo = (f'Planilhas lidas: {len(arquivos)}\nPlanilhas incluídas: {planilhas_incluidas}\nLançamentos incluídos: {incluidos}\n'
              f'Sem lançamento: {ignorados}\nArquivos com erro: {erros}\n')
    resumo += f'\nSaída: {saida}' if incluidos else '\nNenhuma planilha de saída foi criada.'
    if arquivos_com_erro:
        resumo += '\n\nPLANILHAS COM ERRO\n' + '\n'.join(f'- {nome}' for nome in arquivos_com_erro)
    relatorio.write_text(resumo + '\n\n' + '\n'.join(registros), encoding='utf-8-sig')
    progresso('Relatório salvo. Finalizando…')
    restante = max(0, 8 - (time.monotonic() - inicio))
    espera = time.monotonic()
    while time.monotonic() - inicio < 8:
        time.sleep(0.1)
        avancar(95 + 4 * min(1, (time.monotonic() - espera) / max(restante, 0.1)))
    progresso('Concluído.')
    avancar(100)
    return resumo + f'\n\nRelatório: {relatorio}'


LOGO_UFGD = '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'


def janela():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title('Acerto RSC | Divisão de Pagamento de Pessoal')
    root.geometry('1060x900')
    root.minsize(900, 820)
    bg, branco, tinta, discreto, verde = '#f4f6f3', '#ffffff', '#243529', '#66736a', '#567719'
    root.configure(bg=bg)
    root.option_add('*Font', ('Segoe UI', 10))
    estilo = ttk.Style(root)
    estilo.theme_use('clam')
    estilo.configure('Acao.TButton', font=('Segoe UI', 11, 'bold'), padding=(22, 12),
                     background=verde, foreground=branco, borderwidth=0)
    estilo.map('Acao.TButton', background=[('disabled', '#dce3d5'), ('active', '#456112')],
               foreground=[('disabled', '#7f8c76')])
    estilo.configure('Pasta.TButton', padding=(18, 10), background='#eef3e6', foreground=verde, borderwidth=0)
    estilo.map('Pasta.TButton', background=[('active', '#e2ebd5')])
    estilo.configure('Caminho.TEntry', padding=11, fieldbackground='#f8faf7', foreground=tinta,
                     bordercolor='#dce3d9', lightcolor='#dce3d9', darkcolor='#dce3d9')
    estilo.configure('RSC.Horizontal.TProgressbar', background='#789c28', troughcolor='#e8ede3',
                     borderwidth=0, lightcolor='#789c28', darkcolor='#789c28')

    def label(pai, texto='', tamanho=10, cor=tinta, negrito=False, **kwargs):
        ancestral = pai
        while ancestral is not None and not hasattr(ancestral, '_bloco_gradiente'):
            ancestral = getattr(ancestral, 'master', None)
        if ancestral is not None and pai.cget('bg') == branco:
            from tkinter import font as tkfont
            fonte = tkfont.Font(family='Segoe UI', size=tamanho, weight='bold' if negrito else 'normal')
            variavel = kwargs.pop('textvariable', None)
            ancora = kwargs.pop('anchor', 'center')
            justificar = kwargs.pop('justify', 'center')
            quebra = kwargs.pop('wraplength', 0)
            desenho = tk.Canvas(pai, highlightthickness=0, borderwidth=0, bg=branco)
            desenho._fonte = fonte
            def pintar_texto(*args):
                atual = variavel.get() if variavel is not None else texto
                largura = max((fonte.measure(linha) for linha in atual.split('\n')), default=1) + 4
                if quebra:
                    largura = min(largura, quebra)
                desenho.delete('texto')
                w = desenho.winfo_width() if desenho.winfo_width() > 1 else largura
                x = 2 if ancora == 'w' or justificar == 'left' else w / 2
                item = desenho.create_text(x, 2, text=atual, font=fonte, fill=cor,
                                            anchor='nw' if x == 2 else 'n', justify=justificar,
                                            width=quebra or 0, tags='texto')
                bbox = desenho.bbox(item)
                altura = (bbox[3]-bbox[1]+4) if bbox else fonte.metrics('linespace')+4
                if desenho.winfo_reqwidth() != largura or desenho.winfo_reqheight() != altura:
                    desenho.configure(width=largura, height=altura)
                ancestral._bloco_gradiente()
            desenho._pintar_texto = pintar_texto
            desenho.bind('<Configure>', pintar_texto)
            if variavel is not None:
                variavel.trace_add('write', pintar_texto)
            desenho.after_idle(pintar_texto)
            return desenho
        return tk.Label(pai, text=texto, bg=pai.cget('bg'), fg=cor,
                        font=('Segoe UI', tamanho, 'bold' if negrito else 'normal'), **kwargs)

    def bloco_flutuante(pai, expandir=False):
        import math
        tela = tk.Canvas(pai, bg=bg, highlightthickness=0, borderwidth=0, height=210)
        tela.pack(fill='both' if expandir else 'x', expand=expandir, pady=(12, 0))
        conteudo = tk.Frame(tela, bg=branco)
        janela_bloco = tela.create_window(28, 24, window=conteudo, anchor='nw')

        def cor_gradiente(y):
            t = min(1, max(0, (y-3)/max(1, tela.winfo_height()-16)))
            return '#%02x%02x%02x' % tuple(round(a+(b-a)*t) for a,b in zip((255,255,255),(237,244,231)))

        fundos = {}
        pintura_pendente = False
        def pintar_interiores():
            nonlocal pintura_pendente
            pintura_pendente = False
            def visitar(widget):
                if isinstance(widget, tk.Frame) and widget.cget('bg') == branco:
                    if widget not in fundos:
                        fundo = tk.Canvas(widget, highlightthickness=0, borderwidth=0)
                        fundo.place(x=0, y=0, relwidth=1, relheight=1)
                        fundo.tk.call('lower', fundo._w)
                        fundos[widget] = fundo
                    alvo = fundos[widget]
                elif hasattr(widget, '_pintar_texto'):
                    alvo = widget
                else:
                    return
                alvo.delete('gradiente')
                offset = widget.winfo_rooty() - tela.winfo_rooty()
                for y in range(widget.winfo_height()):
                    alvo.create_line(0, y, widget.winfo_width(), y, fill=cor_gradiente(offset+y), tags='gradiente')
                alvo.tag_lower('gradiente')
                for filho in widget.winfo_children():
                    if filho is not fundos.get(widget):
                        visitar(filho)
            visitar(conteudo)

        def solicitar_pintura():
            nonlocal pintura_pendente
            if not pintura_pendente:
                pintura_pendente = True
                tela.after_idle(pintar_interiores)
        conteudo._bloco_gradiente = solicitar_pintura

        def arredondado(x1, y1, x2, y2, raio, cor):
            pontos = [x1+raio,y1,x2-raio,y1,x2,y1,x2,y1+raio,x2,y2-raio,
                      x2,y2,x2-raio,y2,x1+raio,y2,x1,y2,x1,y2-raio,x1,y1+raio,x1,y1]
            tela.create_polygon(pontos, smooth=True, splinesteps=24, fill=cor, outline='', tags='fundo')

        def desenhar(evento=None):
            largura, altura = tela.winfo_width(), tela.winfo_height()
            if largura < 60:
                return
            tela.delete('fundo')
            for margem, cor in [(0, '#edf0eb'), (2, '#e6eae3'), (4, '#dee4da')]:
                arredondado(4+margem, 10+margem, largura-4-margem, altura-2, 25, cor)
            x1, x2, y1, y2, raio = 9, largura-9, 3, altura-13, 23
            for y in range(y1, max(y1+1, y2)):
                t = (y-y1) / max(1, y2-y1)
                # Gradiente pérola: branco no topo, verde muito suave na base.
                rgb = [round(a+(b-a)*t) for a,b in zip((255,255,255),(237,244,231))]
                cor = '#%02x%02x%02x' % tuple(rgb)
                distancia = max(0, raio-(y-y1), raio-(y2-1-y))
                recuo = raio-math.sqrt(max(0, raio*raio-distancia*distancia))
                tela.create_line(x1+recuo, y, x2-recuo, y, fill=cor, tags='fundo')
            tela.tag_lower('fundo')
            tela.itemconfigure(janela_bloco, width=max(1, largura-56))
            if expandir:
                tela.itemconfigure(janela_bloco, height=max(1, altura-58))
            solicitar_pintura()

        def ajustar(evento=None):
            solicitar_pintura()
            if not expandir:
                desejada = conteudo.winfo_reqheight()+58
                if int(float(tela.cget('height'))) != desejada:
                    tela.configure(height=desejada)
        tela.bind('<Configure>', desenhar)
        conteudo.bind('<Configure>', ajustar)
        return conteudo

    def mostrar_conclusao(resumo):
        dialogo = tk.Toplevel(root)
        dialogo.withdraw()
        dialogo.title('Execução concluída')
        dialogo.transient(root)
        dialogo.configure(bg=branco)
        dialogo.resizable(False, False)
        painel_final = tk.Frame(dialogo, bg=branco, padx=38, pady=28)
        painel_final.pack(fill='both', expand=True)
        label(painel_final, '✓', 36, verde, True).pack()
        label(painel_final, 'Processamento concluído', 20, tinta, True).pack(pady=(8, 12))
        dados = resumo.split('\n\n', 1)[0]
        label(painel_final, dados, 11, discreto, justify='center').pack()
        label(painel_final, 'Consulte o relatório de execução para conferir os detalhes.',
              10, discreto, wraplength=470).pack(pady=(16, 22))
        def confirmar(evento=None):
            dialogo.grab_release()
            dialogo.destroy()
        botao_ok = ttk.Button(painel_final, text='OK', style='Acao.TButton', command=confirmar)
        botao_ok.pack(ipadx=40)
        dialogo.update_idletasks()
        w, h = dialogo.winfo_reqwidth(), dialogo.winfo_reqheight()
        x = root.winfo_rootx() + (root.winfo_width()-w)//2
        y = root.winfo_rooty() + (root.winfo_height()-h)//2
        dialogo.geometry(f'{w}x{h}+{max(0,x)}+{max(0,y)}')
        dialogo.protocol('WM_DELETE_WINDOW', confirmar)
        dialogo.bind('<Return>', confirmar)
        dialogo.bind('<Escape>', confirmar)
        dialogo.deiconify()
        dialogo.grab_set()
        botao_ok.focus_set()

    tk.Frame(root, bg='#aec71b', height=4).pack(fill='x')
    cabecalho = tk.Frame(root, bg=branco, padx=30, pady=8)
    cabecalho.pack(fill='x')
    cabecalho.columnconfigure(1, weight=1)
    logo = tk.PhotoImage(data=LOGO_UFGD).subsample(4, 4)
    root.logo_ufgd = logo
    tk.Label(cabecalho, image=logo, bg=branco, borderwidth=0).grid(row=0, column=0, rowspan=2, padx=(0, 22))
    citacao = tk.Frame(cabecalho, bg=branco)
    citacao.grid(row=1, column=1, sticky='se', pady=(10, 6))
    tk.Label(citacao, text='“Quando encontrar uma ideia na qual acredita de todo o coração,\ntrabalhe para realizá-la.” — Henry Ford',
             bg=branco, fg='#788176', font=('Segoe UI', 8, 'italic'),
             justify='right').pack(anchor='e')
    marca = tk.Frame(cabecalho, bg=branco)
    marca.grid(row=0, column=1, sticky='ew', pady=(12, 0))
    label(marca, 'Coordenadoria de administração e pagamento de pessoal', 10, discreto).pack(anchor='w')
    label(marca, 'Divisão de Pagamento de Pessoal', 23, tinta, True).pack(anchor='w', pady=(5, 0))
    tk.Frame(root, bg='#e2e7de', height=1).pack(fill='x')

    corpo = tk.Frame(root, bg=bg, padx=32, pady=20)
    corpo.pack(fill='both', expand=True)
    titulo = tk.Frame(corpo, bg=bg)
    titulo.pack(fill='x', pady=(0, 17))
    label(titulo, 'Acerto RSC', 21, tinta, True).pack(anchor='center')
    label(titulo, 'Identifica todos os lançamentos e os reúne em um único arquivo para macro', 11, discreto).pack(anchor='center', pady=(4, 0))

    cartao = bloco_flutuante(corpo)
    label(cartao, 'Planilhas de origem', 12, tinta, True).pack(anchor='w')
    label(cartao, 'Selecione a pasta que contém os arquivos ODS.', 10, discreto).pack(anchor='w', pady=(3, 12))
    caminho = tk.Frame(cartao, bg=branco)
    caminho.pack(fill='x')
    pasta = tk.StringVar()
    campo_pasta = ttk.Entry(caminho, textvariable=pasta, style='Caminho.TEntry')
    campo_pasta.pack(side='left', fill='x', expand=True, padx=(0, 10))

    def selecionar_pasta():
        escolhida = filedialog.askdirectory(parent=root, title='Selecione a pasta com as planilhas')
        if escolhida:
            pasta.set(escolhida)

    botao_pasta = ttk.Button(caminho, text='Selecionar pasta', command=selecionar_pasta, style='Pasta.TButton')
    botao_pasta.pack(side='right')
    rodape_cartao = tk.Frame(cartao, bg=branco)
    rodape_cartao.pack(fill='x', pady=(16, 0))
    label(rodape_cartao, 'Arquivos .ods  •  Primeira aba de cada arquivo\nSaída em Lançamentos RSC - mês ano',
          9, discreto, justify='left').pack(side='left')

    eventos = queue.Queue()
    ocupado = False
    progresso_real = 0.0
    progresso_visual = 0.0
    ultimo_quadro = time.monotonic()
    conclusao_pendente = None
    status = tk.StringVar(value='Pronto para iniciar')
    porcentagem = tk.StringVar(value='0%')

    def iniciar():
        nonlocal ocupado, progresso_real, progresso_visual, ultimo_quadro, conclusao_pendente
        origem = pasta.get().strip()
        if not origem:
            messagebox.showerror('Selecione uma pasta', 'Escolha a pasta com as planilhas para continuar.', parent=root)
            return
        ocupado = True
        progresso_real = progresso_visual = 0.0
        ultimo_quadro = time.monotonic()
        conclusao_pendente = None
        for controle in controles:
            controle.configure(state='disabled')
        botao_executar.configure(text='Processando…')
        status.set('Iniciando processamento…')
        barra['value'] = 0
        porcentagem.set('0%')
        terminal.configure(state='normal')
        terminal.delete('1.0', 'end')
        terminal.configure(state='disabled')

        def trabalhar():
            try:
                resultado = processar_pasta(origem, lambda msg: eventos.put(('progresso', msg)),
                                            lambda v: eventos.put(('percentual', v)))
                eventos.put(('concluido', resultado))
            except Exception as exc:
                eventos.put(('erro', str(exc)))
        threading.Thread(target=trabalhar, daemon=True).start()

    botao_executar = ttk.Button(rodape_cartao, text='Processar planilhas  →', command=iniciar, style='Acao.TButton')
    botao_executar.pack(side='right')

    acompanhamento = bloco_flutuante(corpo, expandir=True)
    topo = tk.Frame(acompanhamento, bg=branco)
    topo.pack(fill='x')
    label(topo, 'Acompanhe a execução', 12, tinta, True).pack(side='left')
    label(topo, tamanho=11, cor=verde, negrito=True, textvariable=porcentagem).pack(side='right')
    barra = ttk.Progressbar(acompanhamento, maximum=100, mode='determinate', style='RSC.Horizontal.TProgressbar')
    barra.pack(fill='x', pady=(12, 8), ipady=1)
    label(acompanhamento, tamanho=9, cor=discreto, textvariable=status, anchor='w', wraplength=900).pack(fill='x', pady=(0, 12))

    caixa = tk.Frame(acompanhamento, bg='#f1f4f1')
    caixa.pack(fill='both', expand=True)
    topo_terminal = tk.Frame(caixa, bg='#e4ebe1', padx=12, pady=9)
    topo_terminal.pack(fill='x')
    label(topo_terminal, '●  ●  ●', 9, '#779659').pack(side='left')
    label(topo_terminal, 'REGISTRO DE EXECUÇÃO', 8, '#3f5d38').pack(side='left', padx=14)
    terminal = tk.Text(caixa, height=11, bg='#f1f4f1', fg='#36583d', font=('Consolas', 10),
                       relief='flat', borderwidth=0, padx=16, pady=12, wrap='word', state='disabled',
                       spacing1=2, spacing3=3, selectbackground='#d2e4c8', selectforeground='#213d26')
    rolagem = ttk.Scrollbar(caixa, orient='vertical', command=terminal.yview)
    rolagem.pack(side='right', fill='y')
    terminal.configure(yscrollcommand=rolagem.set)
    terminal.pack(fill='both', expand=True)
    terminal.tag_configure('erro', foreground='#a63832')
    terminal.tag_configure('ok', foreground='#36583d')
    terminal.tag_configure('sucesso', foreground='#426b16')

    def registrar(mensagem):
        terminal.configure(state='normal')
        tag = 'erro' if 'ERRO:' in mensagem else ('sucesso' if mensagem.startswith(('INCLUÍDO', 'Concluído')) else 'ok')
        terminal.insert('end', datetime.now().strftime('%H:%M:%S') + '  › ' + mensagem + '\n', tag)
        terminal.see('end')
        terminal.configure(state='disabled')

    registrar('Aguardando a seleção da pasta de origem.')
    label(corpo, 'UFGD  /  Divisão de Pagamento de Pessoal  •  progesp.dpp@ufgd.edu.br', 9, discreto).pack(anchor='w', pady=(12, 0))
    controles = (campo_pasta, botao_pasta, botao_executar)

    def acompanhar():
        nonlocal ocupado, progresso_real, progresso_visual, ultimo_quadro, conclusao_pendente
        try:
            while True:
                tipo, mensagem = eventos.get_nowait()
                if tipo == 'percentual':
                    progresso_real = max(progresso_real, float(mensagem))
                elif tipo == 'progresso':
                    status.set(mensagem)
                    registrar(mensagem)
                else:
                    if tipo == 'concluido' and progresso_visual < 100:
                        conclusao_pendente = (tipo, mensagem)
                        progresso_real = 100.0
                        continue
                    ocupado = False
                    for controle in controles:
                        controle.configure(state='normal')
                    botao_executar.configure(text='Processar planilhas  →')
                    if tipo == 'erro':
                        registrar('ERRO: ' + mensagem)
                        status.set('Não foi possível concluir. Consulte o registro abaixo.')
                        messagebox.showerror('Erro no processamento', mensagem, parent=root)
                    else:
                        status.set('Processamento concluído. Confira os resultados e o relatório.')
                        registrar(mensagem)
                        mostrar_conclusao(mensagem)
        except queue.Empty:
            pass
        agora = time.monotonic()
        decorrido = min(agora - ultimo_quadro, 0.1)
        ultimo_quadro = agora
        if ocupado:
            # Anima a posição entre as atualizações reais, em quadros de 30 ms.
            # O limite de 100% só fica disponível depois de concluir o lote.
            limite = 100.0 if conclusao_pendente else 99.0
            velocidade = 12.5 if progresso_real > progresso_visual else max(0.02, (99 - progresso_visual) / 20)
            progresso_visual = min(limite, progresso_visual + velocidade * decorrido)
            barra['value'] = progresso_visual
            porcentagem.set(f'{int(progresso_visual)}%')
            if conclusao_pendente and progresso_visual >= 100:
                eventos.put(conclusao_pendente)
                conclusao_pendente = None
        root.after(30, acompanhar)

    def fechar():
        if ocupado:
            messagebox.showinfo('Processamento em andamento', 'Aguarde o término da leitura das planilhas.', parent=root)
        else:
            root.destroy()
    root.protocol('WM_DELETE_WINDOW', fechar)
    root.after(100, acompanhar)
    root.mainloop()
    return 0


def main():
    if len(sys.argv) == 1:
        return janela()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--pasta', type=Path, required=True, help='Pasta com arquivos ODS de origem')
    args = parser.parse_args()
    try:
        print(processar_pasta(args.pasta))
    except (OSError, ValueError, KeyError, InvalidOperation) as exc:
        print(f'Erro: {exc}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())



