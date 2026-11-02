# -*- coding: utf-8 -*-
import urllib.parse

def calcular_metricas_thyerri_brl(ads_count: int, days_count: int, cpm: float = 30.0):
    """
    Modelagem Matemática de Leilão (Método B do Thyerri):
    - Estima impressões totais com base no número de criativos ativos e dias no ar.
    - Aplica o CPM médio de R$ 30,00 comum em Low Ticket no Brasil.
    - Projeta o Faturamento Estimado com base no ROAS médio de 1.85x.
    """
    dias_efetivos = min(max(days_count, 5), 60)
    views = max(ads_count * 650 * dias_efetivos, ads_count * 3000)
    gasto = (views / 1000.0) * cpm
    fat = gasto * 1.85
    
    if views < 25000:
        status = "Validação Inicial"
    elif views <= 70000:
        status = "Tração Moderada"
    elif views <= 200000:
        status = "Escalada Lucrativa"
    else:
        status = "Super Escalada (Campeã)"
        
    return {
        "views_estimadas": int(views),
        "gasto_estimado_brl": round(gasto, 2),
        "faturamento_estimado_brl": round(fat, 2),
        "status_thyerri": status
    }

def calcular_metricas_thyerri_usd(ads_count: int, days_count: int, cpm_usd: float = 3.0):
    """
    Adaptação para mercado internacional / LATAM:
    - CPM médio internacional: $3.00 USD.
    - ROAS médio estimado: 2.0x.
    """
    dias_efetivos = min(max(days_count, 5), 60)
    views = max(ads_count * 600 * dias_efetivos, ads_count * 2500)
    gasto = (views / 1000.0) * cpm_usd
    fat = gasto * 2.0
    
    if views < 25000:
        status = "Validação Inicial"
    elif views <= 70000:
        status = "Tração Moderada"
    elif views <= 200000:
        status = "Escalada Lucrativa"
    else:
        status = "Super Escalada (Campeã)"
        
    return {
        "views_estimadas": int(views),
        "gasto_estimado_usd": round(gasto, 2),
        "faturamento_estimado_usd": round(fat, 2),
        "status_thyerri": status
    }

def auditar_criativo_metodo_a(views_reais: int, cpm: float = 30.0):
    """
    Auditoria Cirúrgica do Criativo Campeão (Método A do Thyerri):
    - Recebe o número exato de views visualizado no player público do vídeo no Facebook/Instagram.
    - Calcula o gasto auditado específico daquele criativo.
    """
    gasto = (views_reais / 1000.0) * cpm
    fat_est = gasto * 1.85
    
    if views_reais < 10000:
        status = "Descartar / Rastejando (< 10k views)"
    elif views_reais <= 50000:
        status = "Sinal Moderado de Validação (20k - 50k views)"
    elif views_reais <= 100000:
        status = "Validado com Tração (50k - 100k views)"
    else:
        status = "Super Escalado & Lucrativo (> 100k views)"
        
    return {
        "views_reais": views_reais,
        "gasto_auditado_brl": round(gasto, 2),
        "faturamento_estimado_brl": round(fat_est, 2),
        "classificacao": status
    }

def gerar_link_busca_google(frase_copy: str):
    """Gera o link de busca no Google para localizar a postagem original do vídeo."""
    query = f'"{frase_copy.strip()}" facebook'
    return f"https://www.google.com/search?q={urllib.parse.quote(query)}"
