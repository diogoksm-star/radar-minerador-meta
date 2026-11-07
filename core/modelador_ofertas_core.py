# -*- coding: utf-8 -*-
"""
CORE DO MODELADOR DE OFERTAS THYERRI (WORKSPACES NO ORCA & LOCAL)
Lê as planilhas de mineração, extrai os dados da oferta original e gera
um Workspace completo com Swipe, Raio-X, Funil/Esteira, Copy da LP,
Roteiros de Criativos (Método A) e Especificação Técnica do Entregável.
"""

import os
import re
import sys
import json
import shutil
import urllib.parse
from datetime import datetime
import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

ORCA_WORKSPACES_ROOT = r"C:\Users\diogo\orca\workspaces\Thyerri"
LOCAL_WORKSPACES_ROOT = r"C:\Users\diogo\Desktop\03_CONHECIMENTO_E_SWIPE\Cursos\Thyerri\OFERTAS_MODELADAS"
MINERACOES_ROOT = r"C:\Users\diogo\Desktop\03_CONHECIMENTO_E_SWIPE\Cursos\Thyerri\MINERACOES"

def sanitize_dirname(name):
    clean = re.sub(r'[\\/*?:"<>|]', "", name)
    clean = re.sub(r'\s+', '_', clean.strip())
    return clean[:45]

def encontrar_arquivos_mineracao(pasta_mineracao=None):
    if pasta_mineracao and os.path.exists(pasta_mineracao):
        pastas = [pasta_mineracao]
    else:
        pastas = []
        for d in sorted(os.listdir(MINERACOES_ROOT), reverse=True):
            p = os.path.join(MINERACOES_ROOT, d)
            if os.path.isdir(p) and not d.startswith("00_"):
                pastas.append(p)

    arquivos_encontrados = []
    for p in pastas:
        for f in os.listdir(p):
            if f.endswith(".xlsx") and not f.startswith("~$"):
                full_p = os.path.join(p, f)
                arquivos_encontrados.append(full_p)
    return arquivos_encontrados

def carregar_dados_da_planilha(filepath):
    """
    Tenta carregar os dados de uma planilha (Painel de Modelagem ou Ofertas Validadas).
    Retorna uma lista de dicionários normalizados com dados da oferta e da modelagem.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    ofertas = []

    # Caso 1: Painel de Modelagem por Adjacência (cabeçalhos na linha 5)
    is_painel = False
    header_idx = -1
    for r_idx, r in enumerate(rows[:10]):
        r_str = " ".join([str(c) for c in r if c])
        if "Nome da Oferta Modelada" in r_str or "Novo Ângulo & Diferencial" in r_str:
            is_painel = True
            header_idx = r_idx
            break

    def safe_get(r, idx):
        return str(r[idx] or "").strip() if len(r) > idx and r[idx] is not None else ""

    if is_painel:
        for r_idx in range(header_idx + 1, len(rows)):
            row = rows[r_idx]
            if not row or not row[0]:
                continue
            id_origem = safe_get(row, 0)
            anunciante = safe_get(row, 1)
            prod_orig = safe_get(row, 2)
            longevidade = safe_get(row, 3)
            preco_front = safe_get(row, 4)
            link_lib = safe_get(row, 5)
            
            # Lado direito (Modelagem)
            nome_modelado = safe_get(row, 9)
            subnicho_adj = safe_get(row, 10)
            angulo_adj = safe_get(row, 11)
            formato_adj = safe_get(row, 12)
            esteira_adj = safe_get(row, 13)

            ofertas.append({
                "id": id_origem,
                "anunciante": anunciante,
                "produto_original": prod_orig,
                "longevidade": longevidade,
                "preco_original": preco_front,
                "link_biblioteca_meta": link_lib,
                "nome_modelado": nome_modelado,
                "subnicho_adj": subnicho_adj,
                "angulo_adj": angulo_adj,
                "formato_adj": formato_adj,
                "esteira_adj": esteira_adj,
                "origem_planilha": filepath,
                "pasta_origem": os.path.dirname(filepath)
            })
        return ofertas

    # Caso 2: Planilha de Ofertas Validadas (19 Colunas)
    for r_idx, row in enumerate(rows[1:], start=2):
        if not row or not row[0]:
            continue
        def safe_get(r, idx):
            return str(r[idx] or "").strip() if len(r) > idx and r[idx] is not None else ""

        rank = safe_get(row, 0)
        setor = safe_get(row, 1)
        prod = safe_get(row, 2)
        anunciante = safe_get(row, 3)
        ads_count = safe_get(row, 4)
        longevidade = safe_get(row, 5)
        views = safe_get(row, 6)
        gasto = safe_get(row, 7)
        fat = safe_get(row, 8)
        status_tracao = safe_get(row, 9)
        preco = safe_get(row, 10)
        checkout = safe_get(row, 11)
        lp_url = safe_get(row, 12)
        headline = safe_get(row, 13)
        gancho = safe_get(row, 14)
        link_lib = safe_get(row, 15)
        views_player = safe_get(row, 16)

        ofertas.append({
            "id": rank,
            "setor": setor,
            "anunciante": anunciante,
            "produto_original": prod,
            "ads_count": ads_count,
            "longevidade": longevidade,
            "views_estimadas": views,
            "gasto_estimado": gasto,
            "fat_estimado": fat,
            "status_tracao": status_tracao,
            "preco_original": preco,
            "checkout": checkout,
            "destination_lp": lp_url,
            "headline": headline,
            "gancho_metodo_a": gancho,
            "link_biblioteca_meta": link_lib,
            "views_player": views_player,
            "origem_planilha": filepath,
            "pasta_origem": os.path.dirname(filepath)
        })

    return ofertas

def localizar_oferta(termo_busca, pasta_especifica=None):
    """
    Localiza uma oferta em todas as planilhas ou em uma pasta específica.
    """
    arquivos = encontrar_arquivos_mineracao(pasta_especifica)
    termo_clean = termo_busca.lower().strip().replace("#", "").replace("m-", "")
    
    todas_ofertas = []
    for arq in arquivos:
        try:
            ofs = carregar_dados_da_planilha(arq)
            todas_ofertas.extend(ofs)
        except Exception:
            pass

    # Tenta casar primeiro por ID ou Rank exato
    for o in todas_ofertas:
        oid = o.get("id", "").lower().replace("#", "").replace("m-", "").strip()
        if oid == termo_clean or f"00{termo_clean}" == oid or f"0{termo_clean}" == oid:
            return o, todas_ofertas

    # Tenta casar por nome modelado, produto ou anunciante
    for o in todas_ofertas:
        comb = f"{o.get('nome_modelado', '')} {o.get('produto_original', '')} {o.get('anunciante', '')}".lower()
        if termo_clean in comb:
            return o, todas_ofertas

    return None, todas_ofertas

def enriquecer_oferta_se_necessario(oferta, todas_ofertas):
    """
    Se a oferta veio do Painel de Modelagem, enriquece com dados da planilha de 19 colunas (e vice-versa).
    """
    p_origem = oferta.get("pasta_origem", "")
    # Procura screenshot na pasta de origem
    screen_dir = os.path.join(p_origem, "screenshots_lps_concorrentes")
    if os.path.exists(screen_dir):
        pname_clean = re.sub(r'[^a-zA-Z0-9]', '_', oferta.get("anunciante", ""))[:15]
        for f in os.listdir(screen_dir):
            if pname_clean.lower() in f.lower() or f.startswith(f"Oferta_{oferta.get('id','').replace('#','').replace('M-','').lstrip('0')}_"):
                oferta["screenshot_path"] = os.path.join(screen_dir, f)
                break

    # Se faltar headline ou gancho, cruza com outra oferta do mesmo anunciante
    if not oferta.get("headline") or not oferta.get("nome_modelado"):
        for outra in todas_ofertas:
            if outra.get("anunciante", "").lower() == oferta.get("anunciante", "").lower() and outra != oferta:
                for k in ["headline", "destination_lp", "gancho_metodo_a", "views_player", "nome_modelado", "subnicho_adj", "angulo_adj", "formato_adj", "esteira_adj"]:
                    if not oferta.get(k) and outra.get(k):
                        oferta[k] = outra.get(k)
    return oferta

def extrair_partes_esteira(esteira_str, preco_front):
    """
    Quebra a string de esteira em Front, Bump 1, Bump 2 e sugere Upsell.
    Ex: "Front R$ 37,00 | Bump 1: Tabela de Gás R600a R$ 19,90 | Bump 2: Checklist R$ 14,90"
    """
    bump1_nome = "Acelerador de Aplicação Imediata & Checklist de Bolso"
    bump1_preco = "R$ 19,90"
    bump2_nome = "Tabela Completa de Parâmetros e Valores de Referência"
    bump2_preco = "R$ 14,90"
    upsell_nome = "Mega Pack de Automação & Biblioteca Expandida de Diagnósticos"
    upsell_preco = "R$ 67,00"

    if esteira_str:
        partes = [p.strip() for p in esteira_str.split("|")]
        for p in partes:
            if "bump 1:" in p.lower():
                m = re.search(r'bump 1:\s*(.*?)(R\$\s*[\d,.]+)', p, re.IGNORECASE)
                if m:
                    bump1_nome = m.group(1).strip()
                    bump1_preco = m.group(2).strip()
                else:
                    bump1_nome = p.replace("Bump 1:", "").strip()
            elif "bump 2:" in p.lower():
                m = re.search(r'bump 2:\s*(.*?)(R\$\s*[\d,.]+)', p, re.IGNORECASE)
                if m:
                    bump2_nome = m.group(1).strip()
                    bump2_preco = m.group(2).strip()
                else:
                    bump2_nome = p.replace("Bump 2:", "").strip()

    return {
        "front_preco": preco_front or "R$ 37,00",
        "bump1_nome": bump1_nome,
        "bump1_preco": bump1_preco,
        "bump2_nome": bump2_nome,
        "bump2_preco": bump2_preco,
        "upsell_nome": upsell_nome,
        "upsell_preco": upsell_preco
    }

def criar_workspace_oferta(oferta, criar_no_orca=True, criar_espelho_desktop=True):
    """
    Cria a árvore de arquivos completa do Workspace.
    """
    nome_oferta = oferta.get("nome_modelado") or oferta.get("produto_original") or "Oferta_Sem_Nome"
    safe_folder_name = f"{oferta.get('id', 'OFERTA').replace('#', 'OFERTA_').replace('M-', 'OFERTA_')}_{sanitize_dirname(nome_oferta)}"

    caminhos_criados = []

    # 1. Diretório no Orca
    if criar_no_orca:
        orca_dir = os.path.join(ORCA_WORKSPACES_ROOT, safe_folder_name)
        os.makedirs(orca_dir, exist_ok=True)
        caminhos_criados.append(("Orca Workspace", orca_dir))
        work_dir = orca_dir
    else:
        local_dir = os.path.join(LOCAL_WORKSPACES_ROOT, safe_folder_name)
        os.makedirs(local_dir, exist_ok=True)
        caminhos_criados.append(("Local Workspace", local_dir))
        work_dir = local_dir

    # Subpastas
    swipe_dir = os.path.join(work_dir, "00_SWIPE_E_RAIO_X")
    entregavel_dir = os.path.join(work_dir, "04_ESPECIFICACAO_DO_ENTREGAVEL")
    os.makedirs(swipe_dir, exist_ok=True)
    os.makedirs(entregavel_dir, exist_ok=True)

    # Copia screenshot se houver
    if oferta.get("screenshot_path") and os.path.exists(oferta["screenshot_path"]):
        dest_screen = os.path.join(swipe_dir, "screenshot_concorrente.png")
        shutil.copy2(oferta["screenshot_path"], dest_screen)

    # Atalho para Biblioteca Meta (.url)
    lib_url = oferta.get("link_biblioteca_meta")
    if not lib_url and oferta.get("anunciante"):
        lib_url = f"https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country=BR&q={urllib.parse.quote(oferta['anunciante'])}"
    
    url_file_path = os.path.join(swipe_dir, "abrir_biblioteca_meta.url")
    with open(url_file_path, "w", encoding="utf-8") as f:
        f.write(f"[InternetShortcut]\nURL={lib_url}\n")

    # Parsing da esteira
    esteira = extrair_partes_esteira(oferta.get("esteira_adj", ""), oferta.get("preco_original", "R$ 37,00"))

    # Arquivo 00: DOCUMENTO DE CONCEPÇÃO (ACO - A PLANTA DA CASA SEGUNDO THYERRI)
    concepcao_path = os.path.join(work_dir, "00_DOCUMENTO_DE_CONCEPCAO_ACO.md")
    with open(concepcao_path, "w", encoding="utf-8") as f:
        f.write(f"""# 🏛️ DOCUMENTO DE CONCEPÇÃO DE OFERTA (ACO - MÉTODO THYERRI)
## "A PLANTA DA CASA" — A ESTRUTURA ESTRATÉGICA ANTES DE QUALQUER LINHA DE COPY

> **Princípio Central do Thyerri (MD 01 - Aula 01):**  
> *"A concepção é a planta da casa. Não adianta copy agressiva, página bonita ou tráfego pesado se a concepção for confusa. Tráfego em uma concepção ruim só acelera o prejuízo. Primeiro você pensa e decide (Concepção), depois constrói a entrega (Produto), e por último apresenta e vende (Oferta)."*

---

## 🧭 QUADRO COMPARATIVO: DA OFERTA MÃE À NOVA CONCEPÇÃO MODELADA

| Elemento ACO | 1. Oferta de Referência (Mãe) | 2. Nova Oferta Modelada por Adjacência |
| :--- | :--- | :--- |
| **Nome da Oferta** | **{oferta.get('produto_original')}** | **{nome_oferta}** |
| **Anunciante / Origem** | {oferta.get('anunciante')} ({oferta.get('longevidade', 'Validado')} no ar) | Projeto Novo (Sua Marca / Especialista Modelado) |
| **Nicho Geral** | {oferta.get('setor', 'Técnico / Prático')} | {oferta.get('subnicho_adj', 'Nicho de Mercado Prático')} |
| **Subnicho Específico** | {oferta.get('produto_original')} | {oferta.get('subnicho_adj', 'Aplicação Direta')} |
| **Persona Compradora** | Profissional / Iniciante com dor no bolso ou tempo | Operador / Tutor / Praticante que precisa de resultado imediato |
| **Situação Gatilho** | O momento exato em que a dúvida ou falha trava o trabalho | O momento em que o erro custa caro ou gera constrangimento |
| **A Dor Real** | Insegurança, medo de errar, perda de horas testando | Medo de prejuízo, frustração de tentar sem método claro |
| **O Desejo Urgente** | Resposta pronta em menos de 1 minuto | Saber exatamente o que fazer de primeira |
| **Pensamento Interno** | *"{oferta.get('gancho_metodo_a', 'O que eu faço agora sem errar?')}"* | *"Preciso resolver isso agora sem gastar fortuna nem perder meu dia..."* |
| **Resultado Concreto** | Diagnóstico ou execução correta sem retrabalho | Conclusão rápida da tarefa com segurança técnica total |
| **Mecanismo da Solução** | {oferta.get('angulo_adj', 'Apoio prático consultável')} | {oferta.get('angulo_adj', 'Mecanismo de Consulta Imediata')} |
| **Formato de Entrega** | {oferta.get('formato_adj', 'Guia de Bolso / Tabela')} | {oferta.get('formato_adj', 'Tabela Laminada / Micro-App')} |

---

## 🔬 A DISSECAÇÃO DOS 9 ELEMENTOS DA NOVA OFERTA (ACO)

1. **Nicho:** {oferta.get('setor', 'Mercado Geral')}
2. **Subnicho:** {oferta.get('subnicho_adj', 'Aplicação Específica')}
3. **Persona:** Praticante / Técnico / Dono que não tem tempo a perder com teoria longa.
4. **Situação:** Momento exato em que a dúvida aparece na prática (na bancada, no atendimento ou na execução).
5. **Dor Concreta:** Perda de tempo, insegurança de errar, risco de queimar material/peça cara ou passar vergonha.
6. **Desejo Concreto:** Bater o olho e saber em 30 segundos exatamente o que fazer.
7. **Pensamento Interno:** *"Será que estou fazendo o teste certo? Não posso errar isso agora."*
8. **Resultado Final Mensurável:** Tarefa executada com perfeição e sem retrabalho.
9. **Mecanismo da Solução:** {oferta.get('angulo_adj', 'Consulta rápida e visual')}
10. **Formato:** {oferta.get('formato_adj', 'Guia de Bolso / WebApp')}

---

## 🎯 AS DUAS PERGUNTAS-TESTE DO THYERRI (VALIDAÇÃO INEGOCIÁVEL)

### 🧪 Teste 1 (Validação da Concepção):
> *"Essa ideia resolve uma situação concreta que a pessoa reconhece imediatamente na vida dela?"*
* **RESPOSTA:** **SIM.** A dor é palpável e o cliente reconhece a situação de imediato.

### 🧪 Teste 2 (Validação do Produto Low Ticket):
> *"A pessoa consegue consultar a ferramenta e agir em menos de 10 segundos na hora da dúvida?"*
* **RESPOSTA:** **SIM.** O produto foi concebido como ferramenta de apoio prático (tabela/micro-app), sem teorias longas.

---

## 🛡️ ANÁLISE DE ADJACÊNCIA: NÚCLEO INVARIÁVEL VS. SUPERFÍCIE

* **O Núcleo Invariável (Preservado da Referência):** A motivação psicológica de compra e o mecanismo de alívio rápido de dor.
* **A Superfície (Alterada pela Modelagem):** O novo ângulo, novo subnicho vizinho, nova identidade visual e esteira otimizada com 2 order bumps.
""")

    # Arquivo 00: Raio-X do Concorrente
    raio_x_path = os.path.join(swipe_dir, "raio_x_concorrente.md")
    with open(raio_x_path, "w", encoding="utf-8") as f:
        f.write(f"""# 🕵️ RAIO-X DA OFERTA CONCORRENTE & MOTIVAÇÃO PSICOLÓGICA

> **Data de Análise:** {datetime.now().strftime("%d/%m/%Y")}
> **Referência Minerada:** `{oferta.get('anunciante')}`
> **Status de Validação:** {oferta.get('longevidade', 'Validado')} no ar • {oferta.get('status_tracao', 'Escalada Comprovada')}

---

### 1. Métricas do Anunciante Original
* **Anunciante:** {oferta.get('anunciante')}
* **Produto Original:** {oferta.get('produto_original')}
* **Longevidade Real:** {oferta.get('longevidade', 'N/D')}
* **Preço Praticado no Front:** {oferta.get('preco_original', 'R$ 37,00')}
* **Checkout Identificado:** {oferta.get('checkout', 'Kiwify / Hotmart')}
* **URL da Landing Page:** [{oferta.get('destination_lp', 'N/D')}]({oferta.get('destination_lp', '#')})
* **Biblioteca de Anúncios:** [Ver Criativos Ao Vivo]({lib_url})

---

### 2. Ângulo de Ataque & Headline do Concorrente
* **Headline Original:**
  > *"{oferta.get('headline', 'Headline em análise no swipe')}"*
* **Gancho de Retenção do Criativo (Método A):**
  > *"{oferta.get('gancho_metodo_a', 'Gancho auditado do anúncio campeão')}"*
* **Views Auditadas no Criativo Principal:** {oferta.get('views_player', 'N/D')} views

---

### 3. Engenharia Reversa do "Núcleo Invariável" (Por que as Pessoas Compram?)
* **A Dor Raiz:** O comprador não quer se tornar um acadêmico ou especialista teórico. Ele está com o problema na mão **agora** e está com medo de errar, perder dinheiro ou passar vergonha.
* **O Desejo Imediato:** Uma resposta rápida de 30 segundos. Ele quer olhar uma tabela, digitar em uma calculadora ou bater o olho em uma prancha visual e saber exatamente o que fazer.
* **Onde o Concorrente é Vulnerável:**
  1. Muitas vezes a entrega é um PDF longo e desorganizado de 80 páginas que ninguém lê.
  2. O visual é amador ou antiquado.
  3. Não possui uma esteira otimizada com 2 Order Bumps complementares para alavancar o ticket médio.
""")

    # Arquivo 01: Arquitetura e Funnel Economics
    esteira_path = os.path.join(work_dir, "01_ARQUITETURA_E_ESTEIRA.md")
    with open(esteira_path, "w", encoding="utf-8") as f:
        f.write(f"""# 📐 ARQUITETURA DESTEIRA & FUNNEL ECONOMICS

> **Oferta Modelada:** **{nome_oferta}**
> **Subnicho de Adjacência:** `{oferta.get('subnicho_adj', 'Nicho Técnico')}`
> **Formato Tecnológico:** `{oferta.get('formato_adj', 'Tabela de Bolso + Micro-App')}`

---

### 1. Estrutura de Precificação e Produtos da Esteira

| Posição no Funil | Nome do Produto | Formato | Preço Sugerido | Taxa de Conversão Alvo |
| :--- | :--- | :--- | :---: | :---: |
| **Front-end Principal** | **{nome_oferta}** | {oferta.get('formato_adj', 'Guia de Bolso Laminado')} | **{esteira['front_preco']}** | 100% dos compradores |
| **Order Bump 1** | {esteira['bump1_nome']} | Checklist / Tabela Aceleradora | **{esteira['bump1_preco']}** | 45% a 55% |
| **Order Bump 2** | {esteira['bump2_nome']} | Gabarito Técnico / Ficha Pronta | **{esteira['bump2_preco']}** | 30% a 40% |
| **Upsell 1 (1-Click)** | {esteira['upsell_nome']} | Acervo Completo / Automação | **{esteira['upsell_preco']}** | 12% a 18% |

---

### 2. Projeção de LTV e Alavancagem no Checkout (Exemplo para 100 Vendas)

* **100 Vendas do Front-end ({esteira['front_preco']}):** R$ 3.700,00
* **50 Vendas do Bump 1 ({esteira['bump1_preco']}):** R$ 995,00
* **35 Vendas do Bump 2 ({esteira['bump2_preco']}):** R$ 521,50
* **15 Vendas do Upsell 1 ({esteira['upsell_preco']}):** R$ 1.005,00
* **Faturamento Bruto Total:** **R$ 6.221,50**
* **Ticket Médio Real (AOV):** **R$ 62,21** *(Um aumento de +68% sobre o preço do Front-end de R$ 37!)*

---

### 3. Parâmetros de Tráfego Direto (Meta Ads)
* **CPM Médio Estimado:** R$ 25,00 a R$ 35,00
* **CPA Máximo de Breakeven:** R$ 62,00 por venda do Front
* **CPA Alvo para Escala Lucrativa (ROAS 2.0x):** R$ 31,00
""")

    # Arquivo 02: Copy da Landing Page de Alta Conversão
    copy_path = os.path.join(work_dir, "02_COPY_LANDING_PAGE_ALTA_CONVERSAO.md")
    with open(copy_path, "w", encoding="utf-8") as f:
        f.write(f"""# ✍️ COPY DA LANDING PAGE: {nome_oferta.upper()}

> **Formato:** Página de Vendas Direta de Alta Conversão para Micro-Ferramentas de Baixo Ticket
> **Diferencial Competitivo:** Resposta Rápida • Zero enrolação teórica • Pronto no Celular

---

### [SEÇÃO 1: BARRA DE NOTIFICAÇÃO SUPERIOR]
🔴 **ACESSO IMEDIATO LIBERADO:** Versão Atualizada 2026 com Consulta Rápida de Bolso.

---

### [SEÇÃO 2: HERO SECTION (DOBRA 1)]
* **Superheadline:** PARA TÉCNICOS E PROFISSIONAIS QUE NÃO TÊM TEMPO A PERDER
* **Headline Principal:**
  # "{oferta.get('angulo_adj', 'A Resposta Certa na Palma da Sua Mão em Menos de 30 Segundos')}"
* **Subheadline:**
  *Descubra exatamente onde está a falha sem precisar testar componentes às cegas, sem perder horas em esquemas complicados e sem o risco de queimar placas caras.*
* **Elemento Visual:**
  `[MOCKUP 3D: Celular exibindo o WebApp/Tabela de Consulta Aberta + Versão PDF de Bolso]`
* **Botão CTA Principal (Verde Vibrante):**
  👉 **QUERO TER ACESSO À FERRAMENTA AGORA POR APENAS {esteira['front_preco']}**
* **Micro-texto de Confiança:** *Acesso Vitalício Imediato • Download em 1 Clique • Garantia Blindada de 7 Dias*

---

### [SEÇÃO 3: O GRANDE INIMIGO COMUM (A DOR QUE DÓI NO BOLSO)]
Você já passou por isso?
* Pegar um equipamento ou serviço para fazer, passar 3 horas testando peça por peça e no final descobrir que era apenas um detalhe bobo?
* Ficar inseguro na frente do cliente sem saber se a peça realmente está queimada ou se é erro de sensor?
* Gastar dinheiro do próprio bolso comprando placa nova para descobrir que o defeito continuava lá?

**O problema não é sua competência. O problema é que a maioria dos manuais foram feitos para engenheiros de fábrica, não para quem está no campo de batalha.**

---

### [SEÇÃO 4: A SOLUÇÃO DEFINITIVA (O MECANISMO ÚNICO)]
Apresentamos o **{nome_oferta}**:
Não é um curso com 40 horas de enrolação que você nunca vai ter tempo de assistir.
É uma **Ferramenta de Apoio Prático** feita para você consultar em 30 segundos:
1. **Bata o Olho:** Localize a marca, modelo ou sintoma.
2. **Veja o Diagnóstico:** A tabela mostra a resistência exata, a voltagem esperada e o teste definitivo.
3. **Resolva Sem Dúvidas:** Faça o reparo com 100% de certeza.

---

### [SEÇÃO 5: O QUE VOCÊ VAI RECEBER]
* ✅ **Tabela Mestre em Alta Resolução:** Pronta para consultar no celular ou imprimir plastificada para a bancada.
* ✅ **Mecanismo de Busca Rápida:** Encontre qualquer código de erro ou sintoma por palavras-chave.
* ✅ **Guia Anti-Erros Graves:** Os 5 erros mais comuns que queimam peças e como evitá-los.
* ✅ **Bônus Especial #1:** Checklist de Diagnóstico em Domicílio (para nunca esquecer nenhuma ferramenta).
* ✅ **Bônus Especial #2:** Acesso ao Grupo VIP de Atualizações.

---

### [SEÇÃO 6: CHECKOUT & OFERTA IRRESISTÍVEL]
* Preço Normal Fora da Campanha: ~R$ 197,00~
* **Hoje por Apenas: {esteira['front_preco']}** (ou até 4x no cartão)

👉 **[CLIQUE AQUI PARA DESBLOQUEAR SEU ACESSO IMEDIATO]**

---

### [SEÇÃO 7: GARANTIA BLINDADA DE 7 DIAS]
Teste o **{nome_oferta}** na prática. Se no seu primeiro atendimento ou reparo você achar que ele não te economizou pelo menos 2 horas de trabalho, basta mandar uma mensagem que devolvemos 100% do seu dinheiro. Sem perguntas e sem letras miúdas.

---

### [SEÇÃO 8: PERGUNTAS FREQUENTES (FAQ)]
* **1. Como eu recebo a ferramenta?**
  *Imediatamente após a confirmação do pagamento, você recebe os links no seu e-mail e WhatsApp.*
* **2. Funciona no celular?**
  *Sim, foi 100% otimizada para abrir no smartphone direto na oficina ou na casa do cliente.*
* **3. Preciso pagar mensalidade?**
  *Não! O acesso é único e vitalício.*
""")

    # Arquivo 03: Scripts de Criativos Método A
    scripts_path = os.path.join(work_dir, "03_SCRIPTS_CRIATIVOS_METODO_A.md")
    with open(scripts_path, "w", encoding="utf-8") as f:
        f.write(f"""# 🎬 ROTEIROS DE CRIATIVOS: MÉTODO A (META ADS 9:16)

> **Oferta:** {nome_oferta}
> **Objetivo:** Tráfego Direto para Página de Vendas (Retenção nos 3 primeiros segundos)

---

### 🎥 ROTEIRO 1: "O ERRO COMUM QUE QUEIMA O COMPONENTE"
* **Formato:** Vídeo gravado na bancada/oficina com câmera na mão ou tela dividida.
* **Gancho Visual (0 a 3s):** Close nas mãos com o multímetro tocando o componente enquanto faísca ou com o texto na tela: *"NÃO TROQUE A PLACA ANTES DE VER ISSO!"*
* **Áudio (0 a 5s):**
  > *"{oferta.get('gancho_metodo_a', 'Se o aparelho apresentar esse sinal, nunca cometa o erro de condenar a peça mais cara.')}"*
* **Desenvolvimento (5 a 25s):**
  > *"Em mais de 80% dos casos, o defeito não é o motor nem a placa principal. É apenas a resistência de um sensor que descalibrou. Se você testar na escala de 20k e der fora dessa tabela aqui, o reparo custa menos de R$ 15 e você resolve em 10 minutos."*
* **Demonstração da Ferramenta (25 a 40s):**
  > *"(Mostra a tela do celular abrindo o {nome_oferta}). Olha só: você seleciona a marca, o modelo e em 5 segundos já tem o valor exato da resistência e onde testar."*
* **Chamada para Ação - CTA (40 a 50s):**
  > *"Para parar de perder tempo e ter essa tabela sempre de bolso no seu celular por apenas {esteira['front_preco']}, toca no botão 'Saiba Mais' aqui embaixo agora mesmo."*

---

### 🎥 ROTEIRO 2: "COMPARATIVO: JEITO VELHO VS JEITO NOVO"
* **Formato:** Tela dividida verticalmente.
* **Lado Esquerdo (Jeito Velho):** Cara estressado, cheio de papéis e manuais rasgados, testando fios no escuro.
* **Lado Direito (Jeito Novo):** Técnico abrindo o celular, olhando a tabela em 10 segundos e dando joinha.
* **Áudio:**
  > *"De um lado: 3 horas testando peça no escuro, gastando combustível para comprar peça errada. Do outro: abrir o {nome_oferta} no celular, bater o olho no código e acertar de primeira. O link está na bio ou no botão aqui embaixo."*

---

### 🎥 ROTEIRO 3: "A TABELA QUE TODO PROFISSIONAL DEVERIA TER"
* **Formato:** Demonstração estática de print interativo com zoom dramático.
* **Texto na Tela:** *"Salve esse vídeo antes que você perca dinheiro."*
* **Áudio:**
  > *"Se você trabalha com isso e ainda não tem os valores exatos de teste salvos no celular, você está pagando para trabalhar. Toque em 'Saiba Mais' e baixe a tabela completa antes que saia do ar."*
""")

    # Arquivo 04: Especificação do Entregável
    entregavel_path = os.path.join(entregavel_dir, "especificacao_produto.md")
    with open(entregavel_path, "w", encoding="utf-8") as f:
        f.write(f"""# 🛠️ ESPECIFICAÇÃO TÉCNICA DO ENTREGÁVEL

> **Nome do Produto:** **{nome_oferta}**
> **Tipo:** {oferta.get('formato_adj', 'Tabela de Bolso Laminada em PDF + WebApp de Consulta Rápida')}

---

### 1. Se for WebApp / Micro-App (Prompt para Lovable / Bolt / Cursor):
```markdown
Crie uma aplicação web responsiva (Mobile-First) em React + Tailwind CSS chamada "{nome_oferta}".
Objetivo: Permitir que o usuário encontre códigos de erro, diagnósticos e valores de teste em menos de 10 segundos.

Recursos obrigatórios:
1. Barra de busca com autocomplete rápido.
2. Filtros por Categoria / Marca / Fabricante.
3. Cards limpos mostrando:
   - Sintoma Apresentado
   - Causa Raiz Provável
   - Como Testar (Passo a Passo com fotos ilustrativas)
   - Valor de Referência Esperado
4. Modo Escuro (Dark Mode) de alto contraste para leitura em bancada.
5. Botão "Favoritar" para salvar os testes mais frequentes.
```

---

### 2. Se for Tabela / Guia Laminado (Estrutura para Canva / InDesign):
* **Dimensões:** Formato A4 (210 x 297 mm) e Formato Card Smartphone (1080 x 1920 px).
* **Cores Principais:** Fundo Escuro (#0A2540) com realces em Amarelo Alerta (#FFB300) e Verde Confirmação (#00D4B2).
* **Tipografia:** Sans-Serif legível (Inter ou Roboto) tamanho mínimo 11pt para garantir leitura rápida com pouca luz.
""")

    # Arquivo workspace_meta.json
    meta_path = os.path.join(work_dir, "workspace_meta.json")
    meta_data = {
        "id": oferta.get("id"),
        "nome_oferta": nome_oferta,
        "anunciante_original": oferta.get("anunciante"),
        "subnicho_adj": oferta.get("subnicho_adj"),
        "formato_adj": oferta.get("formato_adj"),
        "esteira": esteira,
        "criado_em": datetime.now().isoformat(),
        "status": "Workspace Pronto para Execução"
    }
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta_data, f, indent=2, ensure_ascii=False)

    # 2. Espelho no Desktop se solicitado
    if criar_espelho_desktop and criar_no_orca:
        local_dir = os.path.join(LOCAL_WORKSPACES_ROOT, safe_folder_name)
        os.makedirs(local_dir, exist_ok=True)
        # Cria um atalho .url para o Orca Workspace
        atalho_orca = os.path.join(local_dir, "Abrir no Orca Workspace.url")
        with open(atalho_orca, "w", encoding="utf-8") as f:
            f.write(f"[InternetShortcut]\nURL=file:///{orca_dir.replace('\\', '/')}\n")
        
        # Copia os arquivos chave para consulta imediata local
        shutil.copy2(raio_x_path, os.path.join(local_dir, "00_RAIO_X_CONCORRENTE.md"))
        shutil.copy2(esteira_path, os.path.join(local_dir, "01_ARQUITETURA_E_ESTEIRA.md"))
        shutil.copy2(copy_path, os.path.join(local_dir, "02_COPY_LANDING_PAGE_ALTA_CONVERSAO.md"))
        shutil.copy2(scripts_path, os.path.join(local_dir, "03_SCRIPTS_CRIATIVOS_METODO_A.md"))
        caminhos_criados.append(("Espelho Local Thyerri", local_dir))

    return work_dir, caminhos_criados, meta_data
