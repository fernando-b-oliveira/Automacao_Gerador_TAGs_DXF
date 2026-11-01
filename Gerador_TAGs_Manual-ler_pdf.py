import os
import re
import math
import unicodedata
from difflib import get_close_matches

import openpyxl
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# =========================================================================
# CONFIGURAÇÕES
# =========================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_EXCEL_PADRAO = os.path.join(BASE_DIR, "exemplos", "tabela_tags_referencia.xlsx")
LAYERS_IGNORADAS = ["CARIMBO", "FORMATO", "FOLHA", "COTA", "DEFPOINT", "EIXO"]

# =========================================================================
# FUNÇÕES DE NORMALIZAÇÃO
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
# MOTOR DO AUTOCAD (DXF) - O CÉREBRO DA OPERAÇÃO
# =========================================================================
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
    ruas, codigos, paredes = [], [], []    
    for e in get_todas_entidades(doc.modelspace()):
        try: layer = str(e.dxf.layer).upper()
        except Exception: continue
        if any(ign in layer for ign in LAYERS_IGNORADAS): continue
        
        tipo = e.dxftype()
        if "SETORIZA" in layer:
            if tipo == 'LINE':
                paredes.append(((e.dxf.start.x, e.dxf.start.y), (e.dxf.end.x, e.dxf.end.y)))
            elif tipo in ('LWPOLYLINE', 'POLYLINE'):
                pts = e.get_points('xy') if tipo == 'LWPOLYLINE' else [(v.dxf.location.x, v.dxf.location.y) for v in e.vertices]
                for i in range(len(pts) - 1):
                    paredes.append((pts[i], pts[i+1]))
                if getattr(e, 'closed', False):
                    paredes.append((pts[-1], pts[0]))
                    
        if tipo not in ("TEXT", "MTEXT", "ARCALIGNEDTEXT"): continue
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
    return ruas, codigos, paredes

def buscar_rua_no_dxf(nome_digitado, ruas):
    alvo = normalizar_texto(nome_digitado).replace("RUA ", "").replace("AVENIDA ", "")
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

# =========================================================================
# GERAÇÃO DAS TAGS (EXCEL)
# =========================================================================
def ccw(A, B, C):
    return (C[1] - A[1]) * (B[0] - A[0]) > (B[1] - A[1]) * (C[0] - A[0])

def intersect(A, B, C, D):
    return ccw(A, C, D) != ccw(B, C, D) and ccw(A, B, C) != ccw(A, B, D)

def achar_setor_cercado(listas_de_ruas, codigos, paredes):
    if not codigos or not listas_de_ruas: return None
    
    melhor_setor = None
    menor_pontuacao = float('inf')
    
    for texto_codigo, (cx, cy) in codigos:
        # PONTUACAO BASE (Lógica Original - Distância)
        distancias_ate_ruas = [min(math.hypot(cx - px, cy - py) for _, _, (px, py) in p_rua) for p_rua in listas_de_ruas]
        if not distancias_ate_ruas: continue
        pontuacao_base = max(distancias_ate_ruas) + (sum(distancias_ate_ruas) * 0.05)
        
        # PONTUACAO DE PAREDES (Nova Lógica - Cruzamentos)
        total_cruzamentos = 0
        if paredes:
            for p_rua in listas_de_ruas:
                # Pega o ponto da rua que está mais próximo desse setor para traçar a linha de visão
                best_p = None
                min_d = float('inf')
                for _, _, (px, py) in p_rua:
                    d = math.hypot(cx - px, cy - py)
                    if d < min_d:
                        min_d = d
                        best_p = (px, py)
                        
                cruzamentos = 0
                for w1, w2 in paredes:
                    if intersect(best_p, (cx, cy), w1, w2):
                        cruzamentos += 1
                total_cruzamentos += cruzamentos
                
        # COMBINANDO AS DUAS LOGICAS
        # Cada parede cruzada adiciona uma penalidade moderada (500 pontos)
        # Assim, setores muito distantes não vencem só por terem 0 paredes, 
        # mas entre setores próximos, o que tiver menos paredes (mesma zona) vence.
        pontuacao = pontuacao_base + (total_cruzamentos * 500)
        
        if pontuacao < menor_pontuacao:
            menor_pontuacao = pontuacao
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
    match = get_close_matches(rua_alvo, list(mapa_ruas.keys()), n=1, cutoff=0.65)
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
# INTERFACE GRÁFICA (GUI) E LÓGICA DE AUTO-COMPLETAR
# =========================================================================
class AppTagGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de TAGs (Busca Rápida de Bairro via CAD)")
        self.root.geometry("700x750")
        self.root.configure(padx=20, pady=20)
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=5)
        style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'))

        self.dxf_cache_path = ""
        self.ruas_dxf = []
        self.codigos_dxf = []
        self.paredes_dxf = []

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

        ttk.Label(root, text="2. Dados da Permissão de Trabalho (PT)", style='Header.TLabel').pack(anchor='w', pady=(15,5))
        
        frame_cliente = ttk.Frame(root)
        frame_cliente.pack(fill='x', pady=(0, 5))
        ttk.Label(frame_cliente, text="Nome do Cliente:", width=15, foreground="green").pack(side='left')
        self.entry_cliente = ttk.Entry(frame_cliente, width=55)
        self.entry_cliente.pack(side='left', padx=5)

        ttk.Label(root, text="Digite o nome da rua que você vê no PDF e clique no botão para puxar as vizinhas:", foreground="blue").pack(anchor='w', pady=(5,5))
        
        # 1ª Rua com o Botão Mágico
        frame_r1 = ttk.Frame(root)
        frame_r1.pack(fill='x', pady=2)
        ttk.Label(frame_r1, text="1ª Rua (Principal):", width=15).pack(side='left')
        self.entry_rua1 = ttk.Entry(frame_r1, width=32)
        self.entry_rua1.pack(side='left', padx=5)
        ttk.Button(frame_r1, text="🏘️ Auto-Completar Vizinhança", command=self.autocompletar_bairro).pack(side='left', padx=5)

        self.entries_ruas = [self.entry_rua1]
        
        # Outras ruas
        for i in range(2, 6):
            frame_r = ttk.Frame(root)
            frame_r.pack(fill='x', pady=2)
            lbl_texto = f"{i}ª Rua {'*' if i <=3 else '(Opcional)'}:"
            ttk.Label(frame_r, text=lbl_texto, width=15).pack(side='left')
            entry_r = ttk.Entry(frame_r, width=55)
            entry_r.pack(side='left', padx=5)
            self.entries_ruas.append(entry_r)

        self.btn_processar = ttk.Button(root, text="🚀 Processar Dados e Gerar Todas as TAGs", command=self.processar)
        self.btn_processar.pack(pady=15, fill='x')

        ttk.Label(root, text="3. Console e Resultados", style='Header.TLabel').pack(anchor='w', pady=(0,2))
        self.text_resultado = tk.Text(root, height=10, width=75, font=('Consolas', 10), bg="#f4f4f4")
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

    def carregar_dxf_memoria(self, caminho):
        if self.dxf_cache_path == caminho: return True
        try:
            self.print_msg("Carregando mapa inteiro do AutoCAD na memória... (Aguarde 1 segundo)")
            self.ruas_dxf, self.codigos_dxf, self.paredes_dxf = carregar_dados_dxf(caminho)
            self.dxf_cache_path = caminho
            return True
        except Exception as e:
            self.print_msg(f"Erro ao carregar DXF: {e}")
            return False

    def autocompletar_bairro(self):
        """A Mágica: Usa a rua digitada para buscar as cruzadas/vizinhas no AutoCAD"""
        nome_digitado = self.entry_rua1.get().strip()
        if len(nome_digitado) < 3:
            messagebox.showwarning("Atenção", "Digite pelo menos uma parte do nome da rua principal (Ex: ARQUIMEDES) para buscar o bairro.")
            return
            
        caminho_dxf = self.entry_dxf.get().strip()
        if not caminho_dxf or not os.path.exists(caminho_dxf):
            messagebox.showwarning("Atenção", "Carregue o arquivo DXF (AutoCAD) primeiro!")
            return
            
        # Carrega o DXF se ainda não carregou
        if not self.ruas_dxf:
            if not self.carregar_dxf_memoria(caminho_dxf): return
            
        self.print_msg(f"Procurando '{nome_digitado}' no mapa do AutoCAD...")
        
        # 1. Acha a rua principal no DXF
        matches = buscar_rua_no_dxf(nome_digitado, self.ruas_dxf)
        if not matches:
            self.print_msg(f"⚠️ Não encontrei nenhuma rua parecida com '{nome_digitado}' no AutoCAD.\nVerifique se o nome está correto.")
            return
            
        ancora = matches[0] # Pega o resultado mais certeiro
        
        # 2. Atualiza a primeira caixa com o nome perfeito que está no CAD
        self.entry_rua1.delete(0, tk.END)
        self.entry_rua1.insert(0, ancora[0])
        
        # 3. Puxa as coordenadas X, Y dessa rua para mapear o bairro
        cx, cy = ancora[2]
        vizinhas = []
        for r in self.ruas_dxf:
            if r[1] == ancora[1]: continue # Ignora a própria rua principal
            dist = math.hypot(cx - r[2][0], cy - r[2][1])
            if dist <= 3800: # Raio ideal para pegar as ruas transversais
                vizinhas.append((r[0], dist))
                
        # 4. Ordena as ruas por proximidade (as mais coladas primeiro)
        vizinhas.sort(key=lambda x: x[1])
        
        # 5. Remove duplicadas mantendo a ordem das mais próximas
        vizinhas_unicas = []
        adicionados = set()
        for v, d in vizinhas:
            nome_limpo = normalizar_texto(v)
            if nome_limpo not in adicionados:
                vizinhas_unicas.append(v)
                adicionados.add(nome_limpo)
                
        # 6. Preenche as caixas debaixo sozinhas
        for i in range(1, 5):
            self.entries_ruas[i].delete(0, tk.END)
            if i - 1 < len(vizinhas_unicas):
                self.entries_ruas[i].insert(0, vizinhas_unicas[i-1])
                
        self.print_msg(f"✅ Vizinhança Mapeada com Sucesso!\nRua âncora: {ancora[0]}\nEncontradas {len(vizinhas_unicas)} ruas ao redor dela.")

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
            messagebox.showwarning("Atenção", "O projeto precisa de pelo menos 3 ruas nas caixas para fechar o setor.")
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
                
            setor_encontrado = achar_setor_cercado(ruas_encontradas, self.codigos_dxf, self.paredes_dxf)
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

