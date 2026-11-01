import os
# Vacinas contra falhas de memória no Windows e desativação de API experimental
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
os.environ["FLAGS_enable_pir_api"] = "0"

import sys
import re
import math
import unicodedata
from difflib import get_close_matches
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
import warnings

import ezdxf
import openpyxl
from PIL import ImageGrab, Image, ImageEnhance
import numpy as np

# Silencia avisos inúteis no terminal
warnings.filterwarnings("ignore", category=DeprecationWarning)
logging.getLogger('ppocr').setLevel(logging.ERROR)
logging.getLogger('paddle').setLevel(logging.ERROR)

from paddleocr import PaddleOCR

# =========================================================================
# INICIALIZAÇÃO DA INTELIGÊNCIA ARTIFICIAL (PADDLE OCR)
# =========================================================================
try:
    # A vacina definitiva contra o erro do C++ é enable_mkldnn=False
    ocr_engine = PaddleOCR(use_textline_orientation=True, lang='pt', enable_mkldnn=False)
except Exception as e:
    messagebox.showerror("Erro Crítico no Paddle", f"Falha ao iniciar IA.\nDetalhes: {e}")
    sys.exit(1)

def run_paddle_ocr(pil_img):
    """
    LÓGICA TESSERACT DENTRO DO PADDLE: 
    Lê as caixas do PaddleOCR e as alinha matematicamente.
    Agrupa os textos que estão na mesma altura para que a Tabela seja lida 
    em formato de 'linha reta' perfeita.
    """
    img_np = np.array(pil_img.convert('RGB'))
    
    # Roda a IA do Baidu
    result = ocr_engine.ocr(img_np)
    
    if not result or not result[0]:
        return ""
        
    boxes_and_texts = []
    for line in result[0]:
        if line and len(line) == 2:
            box = line[0]
            text = line[1][0]
            # Pega o centro Y e X da caixa lida
            center_y = sum([p[1] for p in box]) / 4
            center_x = sum([p[0] for p in box]) / 4
            boxes_and_texts.append({'text': text, 'cx': center_x, 'cy': center_y})
            
    # 1º Passo: Organiza de cima para baixo
    boxes_and_texts.sort(key=lambda x: x['cy'])
    
    lines = []
    current_line = []
    last_y = None
    
    # 2º Passo: Agrupa na mesma linha (margem de 20px)
    for item in boxes_and_texts:
        if last_y is None:
            current_line.append(item)
            last_y = item['cy']
        elif abs(item['cy'] - last_y) < 20: 
            current_line.append(item)
        else:
            lines.append(current_line)
            current_line = [item]
            last_y = item['cy']
    if current_line:
        lines.append(current_line)
        
    # 3º Passo: Organiza da esquerda para a direita (X) e junta com espaços
    text_lines = []
    for line in lines:
        line.sort(key=lambda x: x['cx'])
        text_lines.append(" ".join([item['text'] for item in line]))
        
    return "\n".join(text_lines)

# =========================================================================
# CONFIGURAÇÕES E FUNÇÕES BASE
# =========================================================================
LAYERS_IGNORADAS = ["CARIMBO", "FORMATO", "FOLHA", "COTA", "DEFPOINT", "EIXO"]
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_EXCEL_PADRAO = os.path.join(BASE_DIR, "exemplos", "tabela_tags_referencia.xlsx")
# Raio de tolerância (em unidades do CAD) para barrar ruas alucinadas pelo OCR
RAIO_CONTEXTO = 3000

def normalizar_texto(txt):
    """Limpeza restaurada e blindada para remover prefixos e pontuações."""
    txt = str(txt).strip().upper()
    txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("ASCII")
    txt = re.sub(r"[^A-Z0-9\s]", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    txt = re.sub(r"^(RUA|AVENIDA|AV|R|TRAVESSA|ALAMEDA|TV|RODOVIA|ESTRADA|PCA|PRACA)\s+", "", txt)
    return txt.strip()

def classificar_texto(texto):
    txt_limpo = str(texto).strip().upper()
    if not txt_limpo: return None
    if re.search(r"^(RUA|AV|AVENIDA|R\.|TRAVESSA|ALAMEDA|TV\.|TV|RODOVIA|ESTRADA)\b", txt_limpo):
        return "RUA"
    if len(txt_limpo) > 3:
        return "RUA"
    return None

def limpar_nome_cliente(txt):
    txt = unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')
    return txt.upper().strip()

# =========================================================================
# FERRAMENTA DE CAPTURA DE TELA (SNIPPING TOOL)
# =========================================================================
class SnippingWidget:
    def __init__(self, parent):
        self.top = tk.Toplevel(parent)
        self.top.attributes('-fullscreen', True)
        self.top.attributes('-alpha', 0.3)
        self.top.configure(cursor="cross")
        self.top.attributes('-topmost', True)
        
        self.canvas = tk.Canvas(self.top, bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.start_x = None
        self.start_y = None
        self.rect = None
        self.bbox = None

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

    def on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, 1, 1, outline='red', width=3, fill="white")

    def on_drag(self, event):
        self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

    def on_release(self, event):
        self.bbox = (min(self.start_x, event.x), min(self.start_y, event.y),
                     max(self.start_x, event.x), max(self.start_y, event.y))
        self.top.destroy()

# =========================================================================
# LÓGICA DO AUTOCAD (MIN-MAX RESTAURADO) & EXCEL
# =========================================================================
def get_todas_entidades(msp):
    for e in msp:
        if e.dxftype() == 'INSERT':
            try:
                for sub_e in e.virtual_entities(): yield sub_e
            except Exception: pass
        else: yield e

def carregar_dados_dxf(caminho_dxf):
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
            elif classificar_texto(texto) == "RUA":
                ruas.append((texto, normalizar_texto(texto), (ponto[0], ponto[1])))
        except Exception: pass
    return ruas, codigos

def buscar_rua_no_dxf(nome_digitado, ruas):
    alvo = normalizar_texto(nome_digitado)
    if len(alvo) < 2: return []
    exatos = [r for r in ruas if r[1] == alvo]
    if exatos: return exatos
    contendo = [r for r in ruas if alvo in r[1]]
    if contendo: return contendo
    palavras = alvo.split()
    match_palavras = [r for r in ruas if all(p in r[1] for p in palavras)]
    if match_palavras: return match_palavras
    nomes_unicos = list(set(r[1] for r in ruas))
    proximos = get_close_matches(alvo, nomes_unicos, n=5, cutoff=0.5)
    if proximos: return [r for r in ruas if r[1] in proximos]
    return []

def achar_setor_cercado(listas_de_ruas, codigos):
    """ALGORITMO MIN-MAX: Identificação perfeita do setor."""
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
    """CRIA UMA SIGLA NOVA SE A RUA NÃO EXISTIR NO EXCEL"""
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
    """Retorna a Sigla (e um Boolean dizendo se é Nova). Coluna B(1) e C(2)"""
    aba_ruas = wb["TAG's para Ruas"]
    rua_alvo = normalizar_texto(nome_rua)
    
    for row in aba_ruas.iter_rows(values_only=True):
        if len(row) >= 3 and row[1] is not None and row[2] is not None:
            sigla = str(row[1]).strip()
            nome_planilha = normalizar_texto(str(row[2]))
            if nome_planilha == rua_alvo:
                return sigla, False
                
    for row in aba_ruas.iter_rows(values_only=True):
        if len(row) >= 3 and row[1] is not None and row[2] is not None:
            sigla = str(row[1]).strip()
            nome_planilha = normalizar_texto(str(row[2]))
            if len(nome_planilha) >= 4 and (rua_alvo in nome_planilha or nome_planilha in rua_alvo):
                return sigla, False

    mapa_ruas = {normalizar_texto(str(row[2])): str(row[1]).strip() for row in aba_ruas.iter_rows(values_only=True) if len(row) >= 3 and row[1] is not None and row[2] is not None}
    match = get_close_matches(rua_alvo, list(mapa_ruas.keys()), n=1, cutoff=0.65)
    
    if match: 
        return mapa_ruas[match[0]], False
    
    nova_sigla = gerar_sigla_rua_nova(wb, rua_alvo)
    return nova_sigla, True

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

def gerar_opcoes_de_sigla(nome):
    palavras = [p for p in re.split(r'\W+', limpar_nome_cliente(nome)) if p not in ['E', 'DE', 'DA', 'DO', 'DI', 'DAS', 'DOS', 'JR', 'JUNIOR']]
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
    return list(dict.fromkeys([s[:3] for s in siglas if len(s) == 3]))

def gerar_tag_estacao(wb, nome_cliente):
    if not nome_cliente: return "N/A (Cliente não informado)", "N/A"
    siglas_usadas = set()
    if "Estações PEQ PEAD" in wb.sheetnames:
        for row in wb["Estações PEQ PEAD"].iter_rows(values_only=True):
            for cell in row:
                if isinstance(cell, str):
                    match = re.search(r'SISTEMA_INTERNO_2E-([A-Z0-9]{3})-001', cell)
                    if match: siglas_usadas.add(match.group(1))
    
    opcoes = gerar_opcoes_de_sigla(nome_cliente)
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
# INTERFACE GRÁFICA (GUI)
# =========================================================================
class AppTagGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador Automático de TAGs (PaddleOCR)")
        self.root.geometry("700x800")
        self.root.configure(padx=20, pady=20)
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=5)
        style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'))

        self.dxf_cache_path = ""
        self.ruas_dxf = []
        self.codigos_dxf = []
        self.nomes_ruas_dxf = []
        self.nomes_originais_dxf = {}
        self.ruas_coords_dxf = {}

        # Seção 1
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

        # Seção 2
        ttk.Label(root, text="2. Dados da Permissão de Trabalho (PT)", style='Header.TLabel').pack(anchor='w', pady=(20,5))
        
        frame_cliente = ttk.Frame(root)
        frame_cliente.pack(fill='x', pady=(0, 10))
        ttk.Label(frame_cliente, text="Nome do Cliente:", width=15, foreground="green").pack(side='left')
        self.entry_cliente = ttk.Entry(frame_cliente, width=55)
        self.entry_cliente.pack(side='left', padx=5)

        frame_scanner = ttk.Frame(root)
        frame_scanner.pack(fill='x', pady=(5, 10))
        ttk.Label(frame_scanner, text="Extração Automática:").pack(side='left')
        ttk.Button(frame_scanner, text="📸 Recortar PT + Mapa e Extrair", command=self.iniciar_captura).pack(side='left', padx=10)

        ttk.Label(root, text="A primeira rua preenchida será usada para gerar a Sigla principal da Calçada.", foreground="blue").pack(anchor='w', pady=(0,5))
        
        self.entries_ruas = []
        for i in range(1, 6):
            frame_r = ttk.Frame(root)
            frame_r.pack(fill='x', pady=2)
            lbl_texto = f"{i}ª Rua {'*' if i <=3 else '(Opcional)'}:"
            ttk.Label(frame_r, text=lbl_texto, width=15).pack(side='left')
            entry_r = ttk.Entry(frame_r, width=55)
            entry_r.pack(side='left', padx=5)
            self.entries_ruas.append(entry_r)

        self.btn_processar = ttk.Button(root, text="🚀 Processar Dados e Gerar Todas as TAGs", command=self.processar)
        self.btn_processar.pack(pady=20, fill='x')

        ttk.Label(root, text="3. Resultado Final", style='Header.TLabel').pack(anchor='w', pady=(0,5))
        self.text_resultado = tk.Text(root, height=13, width=75, font=('Consolas', 11), bg="#f4f4f4")
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

    def print_msg(self, msg):
        self.text_resultado.delete(1.0, tk.END)
        self.text_resultado.insert(tk.END, msg)
        self.root.update()

    def carregar_dxf_memoria(self, caminho):
        if self.dxf_cache_path == caminho: return True
        try:
            self.print_msg("Carregando dicionário DXF na memória... (Isso será feito apenas uma vez!)")
            self.ruas_dxf, self.codigos_dxf = carregar_dados_dxf(caminho)
            self.dxf_cache_path = caminho
            self.nomes_ruas_dxf = list(set([normalizar_texto(r[1]) for r in self.ruas_dxf])) 
            self.nomes_originais_dxf = {normalizar_texto(r[1]): r[0] for r in self.ruas_dxf} 
            
            self.ruas_coords_dxf = {}
            for _, normalizado, coords in self.ruas_dxf:
                norm = normalizar_texto(normalizado)
                if norm not in self.ruas_coords_dxf:
                    self.ruas_coords_dxf[norm] = []
                self.ruas_coords_dxf[norm].append(coords)
                
            return True
        except Exception as e:
            self.print_msg(f"Erro ao carregar DXF: {e}")
            return False

    def extrair_dados_tabela_cirurgica(self, texto_ocr):
        cliente = ""
        endereco = ""
        
        match_c = re.search(r'CLIENTE\s*[:|]\s*([^\n\r]+)', texto_ocr, re.IGNORECASE)
        if match_c:
            cliente = match_c.group(1).strip()
            cliente = re.split(r'[-|]', cliente)[0].strip()

        match_e = re.search(r'ENDERE[CÇ]O\s*[:|]\s*([^\n\r]+)', texto_ocr, re.IGNORECASE)
        if match_e:
            endereco = match_e.group(1).strip()
            endereco = re.split(r'[,|-]', endereco)[0].strip()
            endereco = re.sub(r'\s+\d+$', '', endereco).strip()
            
        return cliente, endereco

    def iniciar_captura(self):
        caminho_dxf = self.entry_dxf.get().strip()
        if not caminho_dxf:
            messagebox.showwarning("Atenção", "Selecione o arquivo DXF primeiro!\nO sistema precisa dele como dicionário para a autocorreção.")
            return
            
        if not self.carregar_dxf_memoria(caminho_dxf): return

        self.root.withdraw() 
        snipper = SnippingWidget(self.root)
        self.root.wait_window(snipper.top) 
        self.root.deiconify() 

        if snipper.bbox:
            x1, y1, x2, y2 = snipper.bbox
            if abs(x2-x1) > 10 and abs(y2-y1) > 10:
                img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                self.processar_ocr(img)

    def processar_ocr(self, img):
        self.print_msg("Analisando imagem com PaddleOCR... Aguarde.")
        self.root.update()
        
        try:
            # === LEITURA DA TABELA COM PADDLE OCR ===
            img_enhanced = ImageEnhance.Sharpness(img).enhance(2.0)
            txt_tabela = run_paddle_ocr(img_enhanced)
            
            cliente_txt, endereco_limpo = self.extrair_dados_tabela_cirurgica(txt_tabela)
            
            if cliente_txt:
                self.entry_cliente.delete(0, tk.END)
                self.entry_cliente.insert(0, cliente_txt.title())

            rua_principal = None
            if endereco_limpo:
                alvo_end = normalizar_texto(endereco_limpo)
                if len(alvo_end) >= 3:
                    match_dxf = get_close_matches(alvo_end, self.nomes_ruas_dxf, n=1, cutoff=0.4)
                    if match_dxf:
                        rua_principal = self.nomes_originais_dxf[match_dxf[0]]

            pontos_rua_principal = []
            if rua_principal:
                rua_princ_norm = normalizar_texto(rua_principal)
                pontos_rua_principal = self.ruas_coords_dxf.get(rua_princ_norm, [])

            # === LEITURA DO MAPA ===
            img_rot_90 = img.rotate(-90, expand=True)
            img_rot_270 = img.rotate(90, expand=True)
            
            textos_brutos = [
                txt_tabela, 
                run_paddle_ocr(img_rot_90),
                run_paddle_ocr(img_rot_270)
            ]
            
            texto_mapa = "\n".join(textos_brutos)
            ruas_auxiliares = []
            
            # Filtro Gramatical: Só aceita se tiver formato de Rua
            padrao_rua = r"\b(RUA|AV|AVENIDA|ALAMEDA|TRAVESSA|TV|RODOVIA)\b"
            
            for linha in texto_mapa.split('\n'):
                linha = linha.strip().upper()
                if re.search(padrao_rua, linha):
                    linha_limpa = normalizar_texto(linha)
                    if len(linha_limpa) >= 4:
                        match_dxf = get_close_matches(linha_limpa, self.nomes_ruas_dxf, n=1, cutoff=0.6)
                        if match_dxf:
                            nome_perfeito = self.nomes_originais_dxf[match_dxf[0]]
                            
                            # Se for igual à principal, pula
                            if rua_principal and nome_perfeito == rua_principal:
                                continue
                                
                            aceitar_rua = False
                            if not pontos_rua_principal:
                                aceitar_rua = True 
                            else:
                                pontos_candidato = self.ruas_coords_dxf.get(match_dxf[0], [])
                                for p1 in pontos_rua_principal:
                                    for p2 in pontos_candidato:
                                        dist = math.hypot(p1[0] - p2[0], p1[1] - p2[1])
                                        if dist <= RAIO_CONTEXTO:
                                            aceitar_rua = True
                                            break
                                    if aceitar_rua: break

                            if aceitar_rua and nome_perfeito not in ruas_auxiliares:
                                ruas_auxiliares.append(nome_perfeito)

            # --- 3. PREENCHER INTERFACE ---
            for entry in self.entries_ruas: entry.delete(0, tk.END)
            
            idx = 0
            if rua_principal:
                self.entries_ruas[idx].insert(0, rua_principal)
                idx += 1
                
            for rua in ruas_auxiliares:
                if idx < 5:
                    self.entries_ruas[idx].insert(0, rua)
                    idx += 1
            
            if rua_principal:
                msg = f"✅ SUCESSO!\nRua Principal âncora: {rua_principal}\nMais {len(ruas_auxiliares)} ruas de apoio foram extraídas."
            else:
                msg = f"⚠️ ATENÇÃO: Endereço não lido perfeitamente da tabela.\n{len(ruas_auxiliares)} ruas extraídas do mapa (Preencha a 1ª manualmente)."
                
            self.print_msg(msg)

        except Exception as e:
            self.print_msg(f"Erro no OCR: {e}")

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
        self.print_msg("Processando dados...")
        
        try:
            if not self.carregar_dxf_memoria(caminho_dxf):
                raise Exception("Falha ao abrir arquivo do AutoCAD.")
                
            ruas_encontradas = []
            for nome in nomes_ruas:
                enc = buscar_rua_no_dxf(nome, self.ruas_dxf)
                if not enc: raise Exception(f"Rua '{nome}' não encontrada no mapa AutoCAD.")
                ruas_encontradas.append(enc)
                
            setor_encontrado = achar_setor_cercado(ruas_encontradas, self.codigos_dxf)
            if not setor_encontrado: raise Exception("Nenhum setor foi encontrado dentro desse cerco de ruas.")

            wb = openpyxl.load_workbook(caminho_excel, data_only=True)
            
            rua_principal = nomes_ruas[0]
            
            # --- AGORA A FUNÇÃO RETORNA SE A SIGLA É NOVA OU NÃO ---
            sigla_rua, is_nova_sigla = obter_sigla_rua_excel(wb, rua_principal)
            
            tag_calcada = gerar_tag_calcada(wb, sigla_rua, setor_encontrado)
            tag_estacao, sigla_cliente = gerar_tag_estacao(wb, nome_cliente)
            
            aviso_sigla = ""
            if is_nova_sigla:
                aviso_sigla = f"\n\n[!] AVISO: A rua '{rua_principal}' não existia na planilha.\nUma nova sigla ({sigla_rua}) foi GERADA AUTOMATICAMENTE."
            
            resultado = (
                f"🗺️  SETOR IDENTIFICADO: {setor_encontrado}\n"
                f"{'-'*60}\n"
                f"🟩 TAG DE ESTAÇÃO PEQ PEAD:\n"
                f"Cliente: {nome_cliente if nome_cliente else 'N/A'}\n"
                f"Resultado: {tag_estacao}\n\n"
                f"🟦 TAG DE CALÇADA PEAD:\n"
                f"Rua Ref.: {rua_principal} (Sigla Planilha: {sigla_rua})\n"
                f"Resultado: {tag_calcada}"
                f"{aviso_sigla}" 
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

