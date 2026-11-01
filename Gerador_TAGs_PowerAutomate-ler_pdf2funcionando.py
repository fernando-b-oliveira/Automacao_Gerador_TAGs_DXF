import os
import re
import math
import warnings
import unicodedata
import difflib

# Bibliotecas leves e precisas
import pymupdf
import openpyxl
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# =========================================================================
# CONFIGURAÇÕES E AMBIENTE
# =========================================================================
warnings.filterwarnings("ignore", category=DeprecationWarning)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_EXCEL_PADRAO = os.path.join(BASE_DIR, "exemplos", "tabela_tags_referencia.xlsx")

# =========================================================================
# FUNÇÕES DE NORMALIZAÇÃO E LIMPEZA
# =========================================================================
def normalizar_texto(texto):
    if not texto: return ""
    texto = str(texto).upper().strip()
    substituicoes = {
        "Á": "A", "À": "A", "Ã": "A", "Â": "A", "Ä": "A",
        "É": "E", "È": "E", "Ê": "E", "Ë": "E",
        "Í": "I", "Ì": "I", "Î": "I", "Ï": "I",
        "Ó": "O", "Ò": "O", "Õ": "O", "Ô": "O", "Ö": "O",
        "Ú": "U", "Ù": "U", "Û": "U", "Ü": "U", "Ç": "C"
    }
    for antigo, novo in substituicoes.items():
        texto = texto.replace(antigo, novo)
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = re.sub(r"[^A-Z0-9\s]", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()

def limpar_nome_cliente(txt):
    if not txt: return ""
    txt = unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')
    return txt.upper().strip()

# =========================================================================
# MAPA DO AUTOCAD (DXF) - O CÉREBRO ESPACIAL
# =========================================================================
LAYERS_IGNORADAS = ["CARIMBO", "FORMATO", "FOLHA", "COTA", "DEFPOINT", "EIXO"]

def get_todas_entidades(msp):
    for e in msp:
        if e.dxftype() == 'INSERT':
            try:
                for sub_e in e.virtual_entities(): yield sub_e
            except Exception: pass
        else: yield e

def carregar_dados_dxf(caminho_dxf):
    import ezdxf
    doc = ezdxf.readfile(caminho_dxf)
    ruas, codigos = [], []    
    for e in get_todas_entidades(doc.modelspace()):
        tipo = e.dxftype()
        if tipo not in ("TEXT", "MTEXT", "ARCALIGNEDTEXT"): continue
        try: layer = str(e.dxf.layer).upper()
        except Exception: continue
        if any(ign in layer for ign in LAYERS_IGNORADAS): continue
        try:
            texto = (e.plain_text() if tipo == "MTEXT" else e.dxf.text).strip()
            if not texto: continue
            ponto = e.dxf.insert
            if "SETORIZA" in layer:
                codigos.append((texto, (ponto[0], ponto[1])))
            elif re.search(r"^(RUA|AV|AVENIDA|R\.|TRAVESSA|ALAMEDA|TV\.|TV|RODOVIA|ESTRADA)\b", texto.upper()) or len(texto) > 3:
                txt_norm = normalizar_texto(texto).replace("RUA ", "").replace("AVENIDA ", "")
                ruas.append((texto.upper(), txt_norm, (ponto[0], ponto[1])))
        except Exception: pass
    return ruas, codigos

def buscar_rua_no_dxf(nome_digitado, ruas):
    alvo = normalizar_texto(nome_digitado).replace("RUA ", "").replace("AVENIDA ", "")
    if len(alvo) < 2: return []
    
    exatos = [r for r in ruas if r[1] == alvo]
    if exatos: return exatos
    contendo = [r for r in ruas if alvo in r[1]]
    if contendo: return contendo
    
    nomes_unicos = list(set(r[1] for r in ruas))
    proximos = difflib.get_close_matches(alvo, nomes_unicos, n=3, cutoff=0.6)
    if proximos: return [r for r in ruas if r[1] in proximos]
    return []

# =========================================================================
# LÓGICA DO POWER AUTOMATE (E-MAIL + DXF + PDF)
# =========================================================================
def cruzar_email_dxf_pdf(caminho_pdf, rua_email, cliente_email, ruas_dxf):
    try:
        matches_email = buscar_rua_no_dxf(rua_email, ruas_dxf)
        if not matches_email:
            return None, [], f"A rua do e-mail '{rua_email}' não foi encontrada no mapa do AutoCAD."
            
        ancora = matches_email[0]
        
        cx, cy = ancora[2]
        bairro = {r[1]: r for r in ruas_dxf if math.hypot(cx - r[2][0], cy - r[2][1]) <= 4000}
        
        doc = pymupdf.open(caminho_pdf)
        page = doc[0]
        
        textos_pdf = []
        coord_implantacao = None
        texto_carimbo_completo = ""

        dict_pdf = page.get_text("dict")
        if "blocks" in dict_pdf:
            for block in dict_pdf["blocks"]:
                if "lines" in block:
                    for line in block["lines"]:
                        texto_linha = ""
                        cx_linha = cy_linha = 0
                        for span in line["spans"]:
                            texto_linha += span["text"] + " "
                            bbox = span["bbox"]
                            cx_linha = (bbox[0] + bbox[2]) / 2
                            cy_linha = (bbox[1] + bbox[3]) / 2
                        
                        texto_linha = texto_linha.strip().upper()
                        if len(texto_linha) > 2:
                            textos_pdf.append((texto_linha, cx_linha, cy_linha))
                            texto_carimbo_completo += texto_linha + " "

        doc.close()

        rua_principal = None
        match_carimbo = re.search(r'REDE DA (RUA|AV|AVENIDA|ALAMEDA|TRAVESSA)\s+([A-Z\s]+)', texto_carimbo_completo)
        if match_carimbo:
            nome_carimbo = match_carimbo.group(2).split('-')[0].strip()
            nc_limpo = normalizar_texto(nome_carimbo).replace("RUA ", "").replace("AVENIDA ", "")
            
            for n_bairro, d_bairro in bairro.items():
                if n_bairro in nc_limpo or difflib.SequenceMatcher(None, n_bairro, nc_limpo).ratio() > 0.8:
                    rua_principal = d_bairro
                    break

        if not rua_principal:
            for txt, tx, ty in textos_pdf:
                if "IMPLANTA" in txt or "RAMAL" in txt:
                    coord_implantacao = (tx, ty)
                    break
                    
            if coord_implantacao:
                ruas_no_pdf = []
                for txt, tx, ty in textos_pdf:
                    txt_limpo = normalizar_texto(txt).replace("RUA ", "").replace("AVENIDA ", "")
                    for n_bairro, d_bairro in bairro.items():
                        if n_bairro in txt_limpo or (len(n_bairro) > 5 and difflib.SequenceMatcher(None, n_bairro, txt_limpo).ratio() > 0.85):
                            ruas_no_pdf.append({"dxf": d_bairro, "dist_implantacao": math.hypot(tx - coord_implantacao[0], ty - coord_implantacao[1])})
                
                if ruas_no_pdf:
                    ruas_no_pdf.sort(key=lambda x: x["dist_implantacao"])
                    rua_principal = ruas_no_pdf[0]["dxf"]

        if not rua_principal:
            rua_principal = ancora

        auxiliares = []
        for n_bairro, d_bairro in bairro.items():
            if d_bairro[1] != rua_principal[1]:
                auxiliares.append(d_bairro[0])
                
        return rua_principal[0], auxiliares[:4], "Sucesso"

    except Exception as e:
        return None, [], f"Erro interno: {str(e)}"

# =========================================================================
# LÓGICA DA PLANILHA PARA GERAR TAGS
# =========================================================================
def achar_setor_cercado(listas_de_ruas, codigos):
    if not codigos or not listas_de_ruas: return None
    melhor_setor = None
    menor_distancia_maxima = float('inf')
    for texto_codigo, (cx, cy) in codigos:
        distancias_ate_ruas = [min(math.hypot(cx - px, cy - py) for _, _, (px, py) in p_rua) for p_rua in listas_de_ruas]
        pontuacao = max(distancias_ate_ruas) + (sum(distancias_ate_ruas) * 0.05)
        if pontuacao < menor_distancia_maxima:
            menor_distancia_maxima = pontuacao
            melhor_setor = texto_codigo
    return melhor_setor

def gerar_sigla_rua_nova(wb, nome_rua):
    aba_ruas = wb["TAG's para Ruas"]
    siglas_usadas = set()
    for row in aba_ruas.iter_rows(values_only=True):
        if len(row) >= 3 and row[1] is not None:
            siglas_usadas.add(str(row[1]).strip().upper())
            
    palavras = [p for p in re.split(r'\W+', limpar_nome_cliente(nome_rua)) if p not in ['E', 'DE', 'DA', 'DO', 'DI', 'DAS', 'DOS', 'JR', 'JUNIOR', 'DR', 'PROF']]
    siglas = []
    if len(palavras) >= 3:
        siglas.append(palavras[0][0] + palavras[1][0] + palavras[2][0])
        siglas.append((palavras[0][:2] + palavras[1][0]).upper())
    elif len(palavras) == 2:
        if len(palavras[1]) >= 3: siglas.append((palavras[0][0] + palavras[1][0] + palavras[1][2]).upper())
        siglas.append((palavras[0][0] + palavras[1][:2]).upper())
        siglas.append((palavras[0][:2] + palavras[1][0]).upper())
    elif len(palavras) == 1:
        if len(palavras[0]) >= 3:
            siglas.extend([palavras[0][:3], palavras[0][0] + palavras[0][2:4]])
        else: siglas.append(palavras[0].ljust(3, 'X'))
        
    opcoes = list(dict.fromkeys([s[:3] for s in siglas if len(s) == 3]))
    sigla_def = next((s for s in opcoes if s not in siglas_usadas), None)
    if not sigla_def:
        base = opcoes[0][:2] if opcoes else "XX"
        for i in range(1, 10):
            if f"{base}{i}" not in siglas_usadas:
                sigla_def = f"{base}{i}"
                break
    return sigla_def or "XXX"

def obter_sigla_rua_excel(wb, nome_rua):
    aba_ruas = wb["TAG's para Ruas"]
    rua_alvo = normalizar_texto(nome_rua)
    for row in aba_ruas.iter_rows(values_only=True):
        if len(row) >= 3 and row[1] is not None and row[2] is not None:
            sigla = str(row[1]).strip()
            nome_planilha = normalizar_texto(str(row[2]))
            if nome_planilha == rua_alvo or (len(nome_planilha) >= 4 and (rua_alvo in nome_planilha or nome_planilha in rua_alvo)):
                return sigla, False
            
    mapa_ruas = {normalizar_texto(str(row[2])): str(row[1]).strip() for row in aba_ruas.iter_rows(values_only=True) if len(row) >= 3 and row[1] is not None and row[2] is not None}
    match = difflib.get_close_matches(rua_alvo, list(mapa_ruas.keys()), n=1, cutoff=0.65)
    if match: return mapa_ruas[match[0]], False
    return gerar_sigla_rua_nova(wb, rua_alvo), True

def gerar_tag_calcada(wb, sigla_rua, setor_bruto):
    setor_formatado = str(setor_bruto).zfill(3)
    prefixo_busca = f"C{sigla_rua}-{setor_formatado}-"
    maior_sequencial = 0
    for nome_aba in wb.sheetnames:
        for row in wb[nome_aba].iter_rows(values_only=True):
            for cell in row:
                if isinstance(cell, str) and cell.startswith(prefixo_busca):
                    try:
                        seq_atual = int(cell.split('-')[2])
                        if seq_atual > maior_sequencial: maior_sequencial = seq_atual
                    except Exception: pass 
    return f"{prefixo_busca}{str(maior_sequencial + 1).zfill(2)}"

def gerar_tag_estacao(wb, nome_cliente):
    if not nome_cliente: return "N/A (Cliente não informado)", "N/A"
    siglas_usadas = set()
    if "Estações PEQ PEAD" in wb.sheetnames:
        for row in wb["Estações PEQ PEAD"].iter_rows(values_only=True):
            for cell in row:
                if isinstance(cell, str):
                    match = re.search(r'SISTEMA_INTERNO_2E-([A-Z0-9]{3})-001', cell)
                    if match: siglas_usadas.add(match.group(1))
    
    palavras = [p for p in re.split(r'\W+', limpar_nome_cliente(nome_cliente)) if p not in ['E', 'DE', 'DA', 'DO', 'DI', 'DAS', 'DOS', 'JR', 'JUNIOR']]
    siglas = []
    if len(palavras) >= 3:
        siglas.append(palavras[0][0] + palavras[1][0] + palavras[2][0])
        siglas.append((palavras[0][:2] + palavras[1][0]).upper())
    elif len(palavras) == 2:
        if len(palavras[1]) >= 3: siglas.append((palavras[0][0] + palavras[1][0] + palavras[1][2]).upper())
        siglas.append((palavras[0][0] + palavras[1][:2]).upper())
        siglas.append((palavras[0][:2] + palavras[1][0]).upper())
    elif len(palavras) == 1:
        if len(palavras[0]) >= 3:
            siglas.extend([palavras[0][:3], palavras[0][0] + palavras[0][2:4]])
        else: siglas.append(palavras[0].ljust(3, 'X'))
    opcoes = list(dict.fromkeys([s[:3] for s in siglas if len(s) == 3]))

    sigla_def = next((s for s in opcoes if s not in siglas_usadas), None)
    if not sigla_def:
        base = opcoes[0][:2] if opcoes else "XX"
        for i in range(1, 10):
            if f"{base}{i}" not in siglas_usadas:
                sigla_def = f"{base}{i}"
                break
    sigla_def = sigla_def or "XXX"
    return f"SISTEMA_INTERNO_2E-{sigla_def}-001-FE01", sigla_def

# =========================================================================
# INTERFACE GRÁFICA (GUI) PREPARADA PARA POWER AUTOMATE
# =========================================================================
class AppTagGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Integração E-mail + DXF + PDF")
        self.root.geometry("700x850")
        self.root.configure(padx=20, pady=20)
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=5)
        style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'))

        self.dxf_cache_path = ""
        self.ruas_dxf = []
        self.codigos_dxf = []

        ttk.Label(root, text="1. Caminho dos Arquivos", style='Header.TLabel').pack(anchor='w', pady=(0,5))
        
        frame_dxf = ttk.Frame(root)
        frame_dxf.pack(fill='x', pady=2)
        ttk.Label(frame_dxf, text="Arquivo DXF:", width=15).pack(side='left')
        self.entry_dxf = ttk.Entry(frame_dxf, width=55)
        self.entry_dxf.pack(side='left', padx=5)
        ttk.Button(frame_dxf, text="Procurar", command=self.buscar_dxf).pack(side='left')

        frame_excel = ttk.Frame(root)
        frame_excel.pack(fill='x', pady=2)
        ttk.Label(frame_excel, text="Planilha Excel:", width=15).pack(side='left')
        self.entry_excel = ttk.Entry(frame_excel, width=55)
        self.entry_excel.insert(0, CAMINHO_EXCEL_PADRAO)
        self.entry_excel.pack(side='left', padx=5)
        ttk.Button(frame_excel, text="Procurar", command=self.buscar_excel).pack(side='left')

        frame_pdf = ttk.Frame(root)
        frame_pdf.pack(fill='x', pady=2)
        ttk.Label(frame_pdf, text="PDF do Projeto:", width=15).pack(side='left')
        self.entry_pdf = ttk.Entry(frame_pdf, width=55)
        self.entry_pdf.pack(side='left', padx=5)
        ttk.Button(frame_pdf, text="Procurar PDF", command=self.buscar_pdf_projeto).pack(side='left')

        ttk.Label(root, text="2. Dados Lidos do E-mail (Gatilho Power Automate)", style='Header.TLabel').pack(anchor='w', pady=(15,5))
        
        frame_cliente = ttk.Frame(root)
        frame_cliente.pack(fill='x', pady=(0, 2))
        ttk.Label(frame_cliente, text="Cliente (E-mail):", width=15, foreground="green").pack(side='left')
        self.entry_cliente = ttk.Entry(frame_cliente, width=55)
        self.entry_cliente.pack(side='left', padx=5)

        frame_rua_email = ttk.Frame(root)
        frame_rua_email.pack(fill='x', pady=(0, 5))
        ttk.Label(frame_rua_email, text="Rua Ref. (E-mail):", width=15, foreground="purple").pack(side='left')
        self.entry_rua_email = ttk.Entry(frame_rua_email, width=32)
        self.entry_rua_email.pack(side='left', padx=5)
        ttk.Button(frame_rua_email, text="⚡ Analisar via E-mail", command=self.analisar_via_email).pack(side='left', padx=5)

        ttk.Label(root, text="As ruas identificadas automaticamente aparecerão aqui:", foreground="blue").pack(anchor='w', pady=(5,5))
        
        self.entries_ruas = []
        for i in range(1, 6):
            frame_r = ttk.Frame(root)
            frame_r.pack(fill='x', pady=2)
            lbl_texto = f"{i}ª Rua {'(Principal)' if i == 1 else '(*)' if i <= 3 else '(Opcional)'}:"
            ttk.Label(frame_r, text=lbl_texto, width=15).pack(side='left')
            entry_r = ttk.Entry(frame_r, width=55)
            entry_r.pack(side='left', padx=5)
            self.entries_ruas.append(entry_r)

        self.btn_processar = ttk.Button(root, text="🚀 Processar Dados e Gerar Todas as TAGs", command=self.processar)
        self.btn_processar.pack(pady=10, fill='x')

        ttk.Label(root, text="3. Console e Resultados", style='Header.TLabel').pack(anchor='w', pady=(0,2))
        self.text_resultado = tk.Text(root, height=8, width=75, font=('Consolas', 10), bg="#f4f4f4")
        self.text_resultado.pack(fill='x')

    def buscar_dxf(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos DXF", "*.dxf")])
        if caminho:
            self.entry_dxf.delete(0, tk.END)
            self.entry_dxf.insert(0, caminho)

    def buscar_excel(self):
        caminho = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx")])
        if caminho:
            self.entry_excel.delete(0, tk.END)
            self.entry_excel.insert(0, caminho)

    def buscar_pdf_projeto(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        if caminho:
            self.entry_pdf.delete(0, tk.END)
            self.entry_pdf.insert(0, caminho)

    def carregar_dxf_memoria(self, caminho):
        if self.dxf_cache_path == caminho: return True
        try:
            self.print_msg("Carregando dicionário do AutoCAD na memória...")
            self.ruas_dxf, self.codigos_dxf = carregar_dados_dxf(caminho)
            self.dxf_cache_path = caminho
            return True
        except Exception as e:
            self.print_msg(f"Erro ao carregar DXF: {e}")
            return False

    def analisar_via_email(self):
        caminho_pdf = self.entry_pdf.get().strip()
        rua_email = self.entry_rua_email.get().strip()
        cliente_email = self.entry_cliente.get().strip()
        caminho_dxf = self.entry_dxf.get().strip()

        if not caminho_pdf or not rua_email or not caminho_dxf:
            messagebox.showwarning("Atenção", "Preencha o DXF, o PDF e a Rua do E-mail para executar essa função!")
            return

        if not self.carregar_dxf_memoria(caminho_dxf): return
        
        self.print_msg(f"Iniciando Fluxo Power Automate...\nProcurando Rua '{rua_email}' no DXF e cruzando com o PDF...")
        
        rua_prin, ruas_aux, info = cruzar_email_dxf_pdf(caminho_pdf, rua_email, cliente_email, self.ruas_dxf)
        
        for entry in self.entries_ruas: entry.delete(0, tk.END)
        
        if rua_prin:
            self.entries_ruas[0].insert(0, rua_prin)
            for idx, r_aux in enumerate(ruas_aux):
                if idx + 1 < len(self.entries_ruas):
                    self.entries_ruas[idx + 1].insert(0, r_aux)
            
            self.print_msg(f"✅ Bairro resolvido automaticamente via E-mail!\n📍 Rua Principal: {rua_prin}\n🛣️ Vizinhança Mapeada: {len(ruas_aux)} ruas de apoio.")
        else:
            self.print_msg(f"⚠️ O robô falhou na automação.\n\nMotivo:\n{info}")

    def print_msg(self, msg):
        self.text_resultado.delete(1.0, tk.END)
        self.text_resultado.insert(tk.END, msg)
        self.root.update()

    def processar(self):
        caminho_dxf = self.entry_dxf.get().strip()
        caminho_excel = self.entry_excel.get().strip()
        nome_cliente = self.entry_cliente.get().strip()
        nomes_ruas = [e.get().strip() for e in self.entries_ruas if e.get().strip()]

        if not caminho_dxf or not caminho_excel:
            messagebox.showwarning("Atenção", "Por favor, selecione os caminhos do DXF e do Excel.")
            return
        if len(nomes_ruas) < 3:
            messagebox.showwarning("Atenção", "Insira pelo menos 3 ruas para cercar o ramal.")
            return

        self.btn_processar.config(state='disabled', text="Calculando Geometria e Lendo Planilhas...")
        self.print_msg("Processando dados finais...")
        
        try:
            if not self.carregar_dxf_memoria(caminho_dxf):
                raise Exception("Falha ao abrir mapa do AutoCAD.")
                
            ruas_encontradas = []
            for nome in nomes_ruas:
                enc = buscar_rua_no_dxf(nome, self.ruas_dxf)
                if not enc: raise Exception(f"Rua '{nome}' não encontrada no mapa AutoCAD.")
                ruas_encontradas.append(enc)
                
            setor_encontrado = achar_setor_cercado(ruas_encontradas, self.codigos_dxf)
            if not setor_encontrado: raise Exception("Nenhum setor foi encontrado dentro desse cerco de ruas.")

            wb = openpyxl.load_workbook(caminho_excel, data_only=True)
            rua_principal = nomes_ruas[0]
            sigla_rua, is_nova_sigla = obter_sigla_rua_excel(wb, rua_principal)
            
            tag_calcada = gerar_tag_calcada(wb, sigla_rua, setor_encontrado)
            tag_estacao, sigla_cliente = gerar_tag_estacao(wb, nome_cliente)
            
            aviso_sigla = ""
            if is_nova_sigla:
                aviso_sigla = f"\n\n[!] AVISO: A rua '{rua_principal}' não existia na planilha.\nUma nova sigla ({sigla_rua}) foi GERADA AUTOMATICAMENTE."
            
            resultado = (
                f"🗺️  SETOR IDENTIFICADO: {setor_encontrado}\n"
                f"{'-'*50}\n"
                f"🟩 TAG DE ESTAÇÃO PEQ PEAD:\nCliente: {nome_cliente or 'N/A'}\nResultado: {tag_estacao}\n\n"
                f"🟦 TAG DE CALÇADA PEAD:\nRua Ref.: {rua_principal} (Sigla: {sigla_rua})\nResultado: {tag_calcada}{aviso_sigla}"
            )
            self.print_msg(resultado)
        except Exception as e:
            self.print_msg(f"❌ ERRO NO PROCESSO:\n{str(e)}")
        finally:
            self.btn_processar.config(state='normal', text="🚀 Processar Dados e Gerar Todas as TAGs")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppTagGenerator(root)
    root.mainloop()

