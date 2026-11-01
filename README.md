# CAD Tag Generator - OCR & Spatial Matching
> Automação em Python para extração de dados via OCR / PDF e cruzamento com plantas vetoriais AutoCAD (DXF) para geração estruturada de TAGs por setor.

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![CAD](https://img.shields.io/badge/AutoCAD-DXF%20Vector-green.svg)
![OCR](https://img.shields.io/badge/OCR-PaddleOCR%20%7C%20PyMuPDF-orange.svg)
![License](https://img.shields.io/badge/License-MIT-lightgrey.svg)

---

## 📌 Visão Geral
Este projeto automatiza a identificação e anotação padronizada de equipamentos e componentes em projetos de engenharia. Cruzando dados textuais brutos (de Ordens de Serviço em PDF ou imagens) com a geometria analítica de arquivos vetoriais **AutoCAD (.DXF)**, o sistema determina o setor correspondente e gera a nomenclatura padronizada de TAGs sem necessidade de conferência manual.

```
[Ordem de Serviço (PDF / Imagem)] 
             │
             ▼
   [Extração OCR / PyMuPDF] ──────┐
                                   ▼
[Planta CAD Vetorial (.DXF)] ──► [Motor de Matching Espacial] ──► [TAG Padronizada Gerada]
                                   (Ponto em Polígono / Setor)     Ex: VAL-SET02-001
```

## 🚀 Principais Funcionalidades
- **Extração Rápida de Documentos:** Ingestão de ordens de serviço via `PyMuPDF` e `PaddleOCR` com pré-processamento de imagens.
- **Geometria Analítica Topológica:** Leitura vetorial com `ezdxf` sem dependência de instalação do software CAD (execução *headless* leve e rápida).
- **Matching por Setor:** Cruzamento analítico de nomes de logradouros e setores para evitar inconsistências de cadastro.
- **Formatação de Nomenclatura:** Geração dinâmica no padrão corporativo:
  ```
  [TIPO_EQUIPAMENTO]-[SETOR]-[IDENTIFICADOR_SEQUENCIAL]
  ```

## 🛠️ Tecnologias Utilizadas
- **Linguagem:** Python 3.9+
- **Bibliotecas:** `ezdxf`, `pymupdf` (fitz), `openpyxl`, `paddleocr`, `opencv-python`, `Pillow`.

## 📦 Como Utilizar

### 1. Instalação de Dependências
```bash
pip install ezdxf pymupdf openpyxl opencv-python pillow
```

### 2. Execução
```bash
python Gerador_TAGs_OCR-localizador_de_rua.py
```

## 🛡️ Governança & Segurança
Arquivos de projetos e dados sensíveis de produção foram desvinculados deste repositório. Para reproduzir os testes, utilize arquivos DXF sintéticos de exemplo.
