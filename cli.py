# -*- coding: utf-8 -*-
"""
RADAR MINERADOR META ADS - CLI PRINCIPAL
Uso via Linha de Comando ou Menu Interativo
"""

import sys
import os
import argparse
from core.meta_scraper import MetaAdScraper
from core.lp_dissector import LPDissector
from core.thyerri_metrics import (
    calcular_metricas_thyerri_brl,
    auditar_criativo_metodo_a,
    gerar_link_busca_google
)
from core.excel_exporter import ExcelExporter

def menu_interativo():
    print("="*70)
    print("      RADAR MINERADOR META ADS - LOW TICKET & DIRECT RESPONSE")
    print("      Engenharia Reversa de Ofertas Escaladas (Método Thyerri)")
    print("="*70)
    print("\nEscolha uma opção:")
    print(" [1] Minerar Ofertas na Meta Ad Library (Por Palavra-Chave)")
    print(" [2] Auditar Criativo Campeão (Método A - Views do Player no Facebook)")
    print(" [3] Calcular Estimativa de Gasto e Faturamento (Fórmula Thyerri)")
    print(" [4] Sair")
    
    opcao = input("\nDigite a opção desejada [1-4]: ").strip()
    
    if opcao == "1":
        query = input("\nDigite a palavra-chave de busca (ex: 'guia visual', 'tabela', 'colorimetria'): ").strip()
        dias_min = input("Filtro de longevidade mínima em dias (recomendado: 20): ").strip()
        dias_min = int(dias_min) if dias_min.isdigit() else 20
        
        print(f"\nIniciando busca por '{query}' com filtro de {dias_min}+ dias...")
        scraper = MetaAdScraper(country="BR")
        anuncios = scraper.scrape_query(query, max_scrolls=4)
        
        # Agrupa por anunciante
        grouped = {}
        for a in anuncios:
            pname = a.get("page_name") or "Desconhecido"
            if pname not in grouped:
                grouped[pname] = {
                    "page_name": pname,
                    "ads_count": 0,
                    "max_days": 0,
                    "destination_lp": "",
                    "copies": []
                }
            grouped[pname]["ads_count"] += a.get("duplications", 1)
            if a.get("dias_no_ar", 0) > grouped[pname]["max_days"]:
                grouped[pname]["max_days"] = a.get("dias_no_ar", 0)
            if not grouped[pname]["destination_lp"] and a.get("destination_lp"):
                grouped[pname]["destination_lp"] = a.get("destination_lp")
            if a.get("copy"):
                grouped[pname]["copies"].append(a["copy"])

        # Filtra por longevidade mínima
        validadas = [p for p in grouped.values() if p["max_days"] >= dias_min and p["destination_lp"]]
        validadas.sort(key=lambda x: (x["ads_count"], x["max_days"]), reverse=True)
        
        print(f"\nEncontradas {len(validadas)} ofertas validadas com {dias_min}+ dias no ar:")
        for idx, o in enumerate(validadas[:10], 1):
            m = calcular_metricas_thyerri_brl(o["ads_count"], o["max_days"])
            o["metricas_thyerri"] = m
            gasto_fmt = f"R$ {m['gasto_estimado_brl']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            print(f" #{idx} | {o['page_name']} | {o['ads_count']} ads | {o['max_days']} dias | Gasto Est: {gasto_fmt}")
            print(f"      LP: {o['destination_lp']}")
            
        salvar = input("\nDeseja dissecar as LPs e exportar para Excel? (s/n): ").strip().lower()
        if salvar == "s" and validadas:
            print("\nDissecando Landing Pages e capturando screenshots...")
            dissector = LPDissector(screenshots_dir="output/screenshots")
            for o in validadas[:10]:
                dados_lp = dissector.dissect(o["destination_lp"], o["page_name"])
                o["headline"] = dados_lp["headline"]
                o["checkout"] = dados_lp["checkout"]
                o["preco"] = dados_lp["preco"]
            
            os.makedirs("output", exist_ok=True)
            out_file = "output/resultado_mineracao.xlsx"
            ExcelExporter.export(validadas[:10], out_file)
            print(f"\n[SUCESSO] Planilha salva em: {out_file}")
            print(f"[SUCESSO] Prints salvos em: output/screenshots/")

    elif opcao == "2":
        frase = input("\nCole a primeira frase da copy ou headline do criativo campeão: ").strip()
        link_google = gerar_link_busca_google(frase)
        print(f"\nLink pronto para localizar o post original no Facebook:")
        print(f" -> {link_google}")
        
        views_str = input("\nDigite o número de visualizações visualizado no player (ex: 145000): ").strip()
        views = int(views_str) if views_str.isdigit() else 50000
        
        auditoria = auditar_criativo_metodo_a(views)
        gasto_fmt = f"R$ {auditoria['gasto_auditado_brl']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        fat_fmt = f"R$ {auditoria['faturamento_estimado_brl']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        print("\n" + "="*50)
        print(f"RESULTADO DA AUDITORIA (MÉTODO A):")
        print(f" Views Auditadas no Vídeo: {auditoria['views_reais']:,}".replace(",", "."))
        print(f" Gasto Real Estimado do Criativo: {gasto_fmt} (a CPM R$ 30)")
        print(f" Faturamento Estimado da Operação: {fat_fmt} (ROAS 1.85x)")
        print(f" Veredito de Escala: {auditoria['classificacao']}")
        print("="*50)

    elif opcao == "3":
        ads_str = input("\nQuantidade de anúncios ativos simultâneos: ").strip()
        dias_str = input("Quantidade de dias no ar (longevidade): ").strip()
        ads = int(ads_str) if ads_str.isdigit() else 10
        dias = int(dias_str) if dias_str.isdigit() else 30
        
        m = calcular_metricas_thyerri_brl(ads, dias)
        gasto_fmt = f"R$ {m['gasto_estimado_brl']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        fat_fmt = f"R$ {m['faturamento_estimado_brl']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        print("\n" + "="*50)
        print(f"MODELAGEM DE LEILÃO (MÉTODO B):")
        print(f" Views Estimadas da Conta: {m['views_estimadas']:,}".replace(",", "."))
        print(f" Gasto Estimado em Tráfego: {gasto_fmt}")
        print(f" Faturamento Estimado do Funil: {fat_fmt}")
        print(f" Status Thyerri: {m['status_thyerri']}")
        print("="*50)

def main():
    parser = argparse.ArgumentParser(description="Radar Minerador Meta Ads - Low Ticket")
    parser.add_argument("--mine", type=str, help="Executar mineração para uma query específica (ex: 'guia visual')")
    parser.add_argument("--days", type=int, default=20, help="Filtro de longevidade mínima em dias")
    parser.add_argument("--audit", type=str, help="Headline ou frase de anúncio para gerar busca do Método A")
    parser.add_argument("--views", type=int, default=0, help="Visualizações reais do vídeo para auditoria")
    
    args = parser.parse_args()
    
    if args.mine:
        scraper = MetaAdScraper(country="BR")
        print(f"Minerando: {args.mine} (filtro {args.days}+ dias)...")
        results = scraper.scrape_query(args.mine)
        print(f"Total de anúncios coletados: {len(results)}")
    elif args.audit:
        link = gerar_link_busca_google(args.audit)
        print(f"Link de busca do post: {link}")
        if args.views > 0:
            auditar_criativo_metodo_a(args.views)
    else:
        menu_interativo()

if __name__ == "__main__":
    main()
