# Consolidador de Notas Fiscais (PDF + XML)

Aplicação Streamlit que lê vários DANFEs (PDF) e/ou XMLs de NF-e de uma vez e
consolida os itens em uma única tabela, com download em **Excel (.xlsx)** ou **CSV**.

| Coluna | Origem no XML | Origem no PDF (DANFE) |
|---|---|---|
| Nota Fiscal de Referência | tags `NFref` + texto de `infCpl` | quadro "Informações Complementares" |
| Nota Fiscal de Emissão | `ide/nNF` | chave de acesso (ou campo "Nº") |
| Código do Material | `det/prod/cProd` | coluna "Código Produto" |
| Quantidade | `det/prod/qCom` | coluna "Quant." |

Colunas extras opcionais (barra lateral): descrição, unidade, chave de acesso, arquivo e origem.

## Instalação e execução

Requer Python 3.10 ou superior.

```bash
cd consolidador_nf
python -m venv .venv
# Windows:      .venv\Scripts\activate
# Linux/macOS:  source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

O navegador abre em http://localhost:8501. Arraste os arquivos para a área de upload.

### Arquivos de exemplo e testes (opcional)

```bash
pip install reportlab pytest
python exemplos/gerar_exemplos.py   # gera XMLs e PDFs fictícios em exemplos/
pytest -q                           # 23 testes
```

## Estrutura

```
consolidador_nf/
├── app.py                    # interface Streamlit
├── nf_extrator/
│   ├── xml_parser.py         # leitura do XML (com ou sem namespace, NFe ou nfeProc)
│   ├── pdf_parser.py         # leitura do DANFE com pdfplumber
│   ├── referencias.py        # identifica a NF de referência no texto complementar
│   ├── processador.py        # orquestração, consolidação, deduplicação, agrupamento
│   ├── exportacao.py         # geração do Excel e do CSV
│   ├── modelos.py            # dataclasses ItemNF / ResultadoArquivo
│   └── utils.py              # números BR, chave de acesso, normalização de texto
├── exemplos/gerar_exemplos.py
└── tests/test_extracao.py
```

## Como a NF de referência é identificada

1. Grupos `NFref` do XML (chave da NF referenciada → número).
2. Chaves de acesso de 44 dígitos (com dígito verificador válido) no texto.
3. Menções como "Ref. NF 123", "NF-e nº 000.012.345", "Nota Fiscal de origem: 456",
   "Devolução das NFs 10, 11 e 12" (várias referências ficam separadas por vírgula).
4. Se não houver NF citada, usa o pedido ("Pedido 4500012345"), se existir.
5. Caso contrário: **N/D**.

O número da própria nota é sempre ignorado nessa busca.

## Opções da barra lateral

- **Ignorar notas repetidas** (ligado por padrão): se a mesma nota vier em mais de um
  arquivo (ex.: PDF e XML), mantém só um — o XML tem prioridade. A comparação usa a chave de acesso.
- **Somar itens repetidos**: soma a quantidade de linhas com a mesma NF e o mesmo código.
- **Mostrar colunas extras**.

## Limitações conhecidas

- **PDF escaneado (imagem)** não tem texto e não é lido — a nota aparece com erro no
  quadro "Status por arquivo". Use o XML (ou um OCR externo antes).
- O leitor de PDF reconhece o **layout padrão do DANFE** (Código, Descrição, NCM, CST,
  CFOP, Un, Quant, Valor...). Layouts muito customizados podem não ter itens reconhecidos;
  nesses casos a aplicação avisa. Códigos de produto contendo espaços também podem falhar no PDF.
- Sempre que possível prefira o XML: é a fonte oficial e a leitura é exata.
