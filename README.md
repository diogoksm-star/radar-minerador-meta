# 🎯 Radar Minerador Meta Ads (Low Ticket & Direct Response)

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![Playwright](https://img.shields.io/badge/Playwright-Automated-green.svg)](https://playwright.dev/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Metodologia](https://img.shields.io/badge/Metodologia-Thyerri%20Low%20Ticket-orange.svg)]()

Ferramenta de engenharia reversa e inteligência competitiva para **minerar, dissecar e auditar ofertas de Low Ticket e Tráfego Direto** na **Meta Ad Library (Biblioteca de Anúncios do Facebook)**.

Desenvolvido para encontrar infoprodutos, micro-apps, guias práticos, tabelas e ferramentas digitais que estão **realmente vendendo no tráfego frio**, eliminando a busca cega e o lixo institucional.

---

## ⚡ Por que esta ferramenta existe?

Se você entrar na Biblioteca de Anúncios do Meta e pesquisar `emagrecimento`, `bíblia` ou `confeitaria`, 95% do resultado será:
* Grandes marcas institucionais;
* Posts impulsionados de influenciadores;
* Anúncios que subiram ontem (0 ou 1 dia) e serão pausados amanhã por darem prejuízo.

O **Radar Minerador** opera sob as regras de **Tração Real**:
1. **Filtro de Longevidade Inegociável:** Filtra anúncios rodando há **20, 30, 60, 90 e até 120+ dias ininterruptos**. Quem gasta dinheiro no Meta Ads por mais de 30 dias com o mesmo criativo está obrigatoriamente no lucro.
2. **Dissecação Automatizada de LPs:** Abre as páginas de vendas dos concorrentes, detecta plataformas de checkout (**Kiwify, Hotmart, Cakto, Eduzz, Cartpanda**), extrai preços de front-end e tira prints em alta resolução.
3. **Cálculo Financeiro de Tráfego (Método B):** Estima as visualizações acumuladas e o **gasto em dinheiro da conta** usando a fórmula de leilão a **CPM médio de R$ 30,00**.
4. **Auditoria do Criativo Campeão (Método A):** Protocolo para localizar o post original do vídeo no Facebook/Instagram e ler as views reais do player.

---

## 🛠️ Arquitetura do Projeto

```
radar-minerador-meta/
│
├── cli.py                        # Terminal Interativo e interface de linha de comando
├── requirements.txt              # Dependências (Playwright, BeautifulSoup4, OpenPyXL)
├── README.md                     # Documentação completa do projeto
├── LICENSE                       # Licença de código aberto MIT
│
├── core/                         # Módulos centrais
│   ├── meta_scraper.py           # Raspagem inteligente na Meta Ad Library (scroll & cards)
│   ├── lp_dissector.py           # Dissecação de Landing Pages, preços, checkouts e prints
│   ├── thyerri_metrics.py        # Fórmulas de tráfego, CPM R$ 30 e auditoria do Método A
│   └── excel_exporter.py         # Gerador de planilhas executivas com paleta executiva
│
├── prompts/                      # Inteligência de Copywriting
│   └── PROMPT_MESTRE_MINERADOR_KEYWORDS.md # Gerador de palavras-chave nos 7 clusters
│
└── examples/                     # Amostras reais de dados minerados
    └── exemplo_ofertas_validadas.xlsx # Planilha pronta com 15 ofertas reais auditadas
```

---

## 🚀 Como Instalar e Rodar

### 1. Clonar o Repositório e Criar Ambiente Virtual
```bash
git clone https://github.com/diogoksm-star/radar-minerador-meta.git
cd radar-minerador-meta

python -m venv venv
# No Windows:
venv\Scripts\activate
# No Linux/Mac:
source venv/bin/activate
```

### 2. Instalar Dependências e Navegadores
```bash
pip install -r requirements.txt
playwright install chromium
```

---

## 💻 Como Usar

### Modo 1: Menu Interativo no Terminal (Recomendado)
Basta executar:
```bash
python cli.py
```
O sistema abrirá o painel:
```text
======================================================================
      RADAR MINERADOR META ADS - LOW TICKET & DIRECT RESPONSE
      Engenharia Reversa de Ofertas Escaladas (Método Thyerri)
======================================================================

Escolha uma opção:
 [1] Minerar Ofertas na Meta Ad Library (Por Palavra-Chave)
 [2] Auditar Criativo Campeão (Método A - Views do Player no Facebook)
 [3] Calcular Estimativa de Gasto e Faturamento (Fórmula Thyerri)
 [4] Sair
```

### Modo 2: Linha de Comando Direta (CLI)

* **Minerar ofertas com mais de 25 dias no ar:**
  ```bash
  python cli.py --mine "tabela de consulta" --days 25
  ```

* **Auditar Criativo Campeão via Método A:**
  ```bash
  python cli.py --audit "PARE DE TROCAR PEÇA BOA" --views 110000
  ```

---

## 📊 A Fórmula de Tração Real (Método Thyerri)

Como o Meta não abre as impressões de anúncios comerciais, aplicamos a engenharia reversa de leilão:

### Método B (Leilão em Lote)
* Um anúncio ativo em escala gasta no mínimo **R$ 15 a R$ 25/dia** por conjunto.
* A um **CPM médio de R$ 30,00**, cada anúncio consome cerca de **650 impressões diárias**.
$$\text{Views Estimadas} = \text{Ads Ativos} \times 650 \text{ imp/dia} \times \text{Dias Úteis (max 60)}$$
$$\text{Gasto Estimado} = \left(\frac{\text{Views}}{1.000}\right) \times \text{R\$\ 30,00}$$

### Método A (A Última Etapa: Auditoria do Criativo Campeão)
Para a oferta vencedora que você escolheu modelar:
1. Copie a primeira frase da copy do criativo mais escalado.
2. Jogue no Google: `"[Primeira frase da copy]" facebook`.
3. Abra o post público do vídeo no Facebook/Instagram.
4. Leia o contador exato de visualizações no player (ex: 172.000 views).
5. O script calcula na hora o investimento real injetado naquele vídeo:
$$\text{Gasto Real do Criativo} = \left(\frac{172.000}{1.000}\right) \times \text{R\$\ 30,00} \approx \mathbf{R\$\ 5.160,00}$$

---

## 📁 Planilha Gerada

A ferramenta exporta uma planilha executiva com **18 colunas** prontas:
* **Identificação:** Rank, Nicho/Setor, Nome do Produto, Anunciante;
* **Métricas Meta:** Ads Ativos Simultâneos, Longevidade Real;
* **Financeiro Global:** Views Estimadas da Conta, Gasto Estimado (R$), Faturamento Projetado (ROAS 1.85x);
* **Dissecação:** Preço Front-End, Plataforma de Checkout, Link da Landing Page, Headline;
* **Auditoria Método A (Colunas Verdes):** Gancho de Busca no Google, Link do Post do Vídeo, Views Reais no Player e Gasto Auditado do Criativo Individual.

---

## 🤝 Contribuições

Contribuições são muito bem-vindas! Se você desenvolveu novos seletores para outros países, novas plataformas de checkout (ex: Stripe, Cartpanda Global) ou novas regras de classificação, sinta-se livre para abrir um **Pull Request** ou uma **Issue**.

---

## 📜 Licença

Distribuído sob a licença **MIT**. Veja [`LICENSE`](LICENSE) para mais detalhes.
