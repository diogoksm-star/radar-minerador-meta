# -*- coding: utf-8 -*-
"""
CLI RUNNER: MODELADOR DE OFERTAS THYERRI
Uso:
  python modelar_oferta.py 1
  python modelar_oferta.py M-001
  python modelar_oferta.py "Geladeiras"
  python modelar_oferta.py --listar
"""

import sys
import os

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modelador_ofertas_core import (
    localizar_oferta,
    enriquecer_oferta_se_necessario,
    criar_workspace_oferta,
    encontrar_arquivos_mineracao,
    carregar_dados_da_planilha
)

def listar_ofertas_disponiveis():
    print("=" * 75)
    print("  OFERTAS DISPONÍVEIS NAS PLANILHAS DE MINERAÇÃO")
    print("=" * 75)
    arquivos = encontrar_arquivos_mineracao()
    total = 0
    for arq in arquivos:
        nome_arq = os.path.basename(arq)
        ofs = carregar_dados_da_planilha(arq)
        if not ofs:
            continue
        print(f"\n📂 Planilha: {nome_arq} ({len(ofs)} ofertas)")
        for o in ofs[:12]:
            total += 1
            nome = o.get("nome_modelado") or o.get("produto_original") or "Oferta"
            orig = o.get("anunciante") or o.get("produto_original") or ""
            print(f"  [{o.get('id', '#')}] {nome[:45]:<45} | Ref: {orig[:25]}")
    print("\n" + "=" * 75)
    print(f"  Total indexado: {total} ofertas.")
    print("  Para modelar uma oferta: python modelar_oferta.py <ID ou Nome>")
    print("=" * 75)

def main():
    if len(sys.argv) < 2 or sys.argv[1] in ["--help", "-h"]:
        print("Uso: python modelar_oferta.py <ID_DA_OFERTA_OU_NOME>")
        print("Exemplos:")
        print("  python modelar_oferta.py 1")
        print("  python modelar_oferta.py M-001")
        print("  python modelar_oferta.py Geladeiras")
        print("  python modelar_oferta.py --listar")
        return

    termo = sys.argv[1].strip()
    if termo == "--listar":
        listar_ofertas_disponiveis()
        return

    print(f"🔍 Buscando oferta com termo: '{termo}' nas planilhas de mineração...")
    oferta, todas = localizar_oferta(termo)

    if not oferta:
        print(f"❌ Nenhuma oferta encontrada com o termo '{termo}'.")
        print("Dica: Use 'python modelar_oferta.py --listar' para ver todas as disponíveis.")
        return

    oferta = enriquecer_oferta_se_necessario(oferta, todas)

    nome_disp = oferta.get('nome_modelado') or oferta.get('produto_original')
    print(f"✅ Oferta localizada: [{oferta.get('id')}] {nome_disp}")
    print(f"   Referência Original: {oferta.get('anunciante')} ({oferta.get('longevidade', 'Validado')})")
    print("🚀 Construindo Workspace de Modelagem...")

    work_dir, caminhos, meta = criar_workspace_oferta(oferta, criar_no_orca=True, criar_espelho_desktop=True)

    print("\n" + "=" * 75)
    print("🎉 WORKSPACE CRIADO COM SUCESSO!")
    print("=" * 75)
    for rotulo, p in caminhos:
        print(f"📁 {rotulo}: {p}")
    print("\n📄 Arquivos gerados:")
    print("  ├── 00_SWIPE_E_RAIO_X/raio_x_concorrente.md")
    print("  ├── 00_SWIPE_E_RAIO_X/abrir_biblioteca_meta.url")
    if oferta.get("screenshot_path"):
        print("  ├── 00_SWIPE_E_RAIO_X/screenshot_concorrente.png")
    print("  ├── 01_ARQUITETURA_E_ESTEIRA.md")
    print("  ├── 02_COPY_LANDING_PAGE_ALTA_CONVERSAO.md")
    print("  ├── 03_SCRIPTS_CRIATIVOS_METODO_A.md")
    print("  ├── 04_ESPECIFICACAO_DO_ENTREGAVEL/especificacao_produto.md")
    print("  └── workspace_meta.json")
    print("=" * 75)

if __name__ == "__main__":
    main()
