# -*- coding: utf-8 -*-
"""
═══════════════════════════════════════════════════════════════════
  OfertasTec - Expansor de Links de Afiliado
═══════════════════════════════════════════════════════════════════

  Este script:
  - Le sua planilha Produtos.xlsx
  - Pega todos os links meli.la
  - Segue o redirect de cada um (usando o IP do seu PC)
  - Descobre se cada link aponta para produto especifico ou vitrine
  - Salva tudo em link_map.json

  Como rodar (PowerShell):
    1. pip install pandas openpyxl requests
    2. cd $env:USERPROFILE\Desktop\OfertasTec
    3. python expandir_links.py

═══════════════════════════════════════════════════════════════════
"""

import json
import re
import sys
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import pandas as pd
    import requests
except ImportError:
    print("[ERRO] Faltam bibliotecas. Rode antes:")
    print("       pip install pandas openpyxl requests")
    sys.exit(1)

PLANILHA = 'Produtos.xlsx'
OUTPUT = 'link_map.json'

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
      "AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/120.0.0.0 Safari/537.36")


def expandir(short_url):
    """Segue redirects e classifica o link."""
    try:
        r = requests.get(
            short_url,
            headers={'User-Agent': UA},
            allow_redirects=True,
            timeout=15
        )
        final = r.url

        # Tenta extrair MLB id
        mlb_match = (
            re.search(r'MLB-?(\d{8,12})', final)
            or re.search(r'/p/MLB(\d{8,12})', final)
        )
        mlb_id = f"MLB{mlb_match.group(1)}" if mlb_match else None

        # Classifica
        if '/social/' in final and not mlb_id:
            kind = 'vitrine_social'
        elif mlb_id and '/p/' in final:
            kind = 'catalogo'
        elif mlb_id:
            kind = 'produto'
        else:
            kind = 'outro'

        return {
            'short': short_url,
            'final': final,
            'kind': kind,
            'mlb_id': mlb_id,
            'status': r.status_code,
        }
    except Exception as e:
        return {
            'short': short_url,
            'final': None,
            'kind': 'erro',
            'mlb_id': None,
            'error': str(e),
        }


def main():
    print("=" * 60)
    print("  OfertasTec - Expansor de Links")
    print("=" * 60)

    if not os.path.exists(PLANILHA):
        print(f"[ERRO] Arquivo '{PLANILHA}' nao encontrado nesta pasta.")
        print(f"       Pasta atual: {os.getcwd()}")
        print(f"       Coloque o Produtos.xlsx aqui e rode de novo.")
        input("\nPressione ENTER para sair...")
        return

    try:
        df = pd.read_excel(PLANILHA, sheet_name=0)
    except Exception as e:
        print(f"[ERRO] Nao consegui ler a planilha: {e}")
        input("\nPressione ENTER para sair...")
        return

    link_to_cat = {}
    for col in df.columns:
        for v in df[col].dropna().astype(str):
            v = v.strip()
            if v.startswith('http'):
                link_to_cat[v] = col

    links = list(link_to_cat.keys())
    print(f"[OK] Encontrados {len(links)} links unicos em {len(df.columns)} colunas")
    print(f"[..] Expandindo (pode levar 2-5 minutos)...\n")

    resultados = []
    with ThreadPoolExecutor(max_workers=6) as ex:
        futures = {ex.submit(expandir, l): l for l in links}
        for i, fut in enumerate(as_completed(futures)):
            r = fut.result()
            r['categoria'] = link_to_cat.get(r['short'], '')
            resultados.append(r)
            mlb = r.get('mlb_id') or ''
            print(f"  [{i+1:3}/{len(links)}] {r['kind']:18} {mlb}")

    print("\n" + "=" * 60)
    print("  RESUMO")
    print("=" * 60)
    from collections import Counter
    kinds = Counter(r['kind'] for r in resultados)
    total = len(resultados)
    for k, v in kinds.most_common():
        pct = v * 100 / total
        print(f"  {k:20} {v:4} ({pct:.1f}%)")

    with open(OUTPUT, 'w', encoding='utf-8') as f:
        json.dump(resultados, f, indent=2, ensure_ascii=False)
    print(f"\n[OK] Salvo em '{OUTPUT}'")

    print("\n" + "=" * 60)
    print("  VEREDITO")
    print("=" * 60)
    bom = kinds.get('produto', 0) + kinds.get('catalogo', 0)
    if bom == total:
        print("TODOS os links apontam pra produto/catalogo especifico!")
        print("Da pra fazer match perfeito de nome/foto/preco.")
    elif bom > total * 0.7:
        print(f"{bom}/{total} links sao produto/catalogo.")
        print("Da pra fazer match perfeito na maioria.")
    elif bom > 0:
        print(f"Apenas {bom}/{total} links sao produto especifico.")
        print("Vale considerar regerar os outros no Gerador do ML.")
    else:
        print("Nenhum link e produto especifico - todos sao vitrine.")
        print("Estrategia atual (busca por categoria) continua a melhor.")

    print(f"\n>> Proximo passo: envie o arquivo '{OUTPUT}' para o Claude!")
    print(f"   Local: {os.path.join(os.getcwd(), OUTPUT)}")
    input("\nPressione ENTER para sair...")


if __name__ == '__main__':
    main()
