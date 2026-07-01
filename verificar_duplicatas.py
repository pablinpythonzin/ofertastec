"""
═══════════════════════════════════════════════════════════════
  VERIFICAR DUPLICATAS — OfertasTec
  Checa a planilha Produtos_ml.xlsx em busca de erros
═══════════════════════════════════════════════════════════════

O QUE VERIFICA:
1. ✅ Links duplicados (mesmo meli.la/XXX em mais de uma linha)
2. ✅ Nomes de produtos duplicados ou muito parecidos
3. ✅ Linhas com link mas sem nome (esquecimento)
4. ✅ Linhas com nome mas sem link (errado)
5. ✅ Categorias fora do padrão (digitação errada)
6. ✅ Links que não são meli.la (suspeitos)
7. ✅ Resumo geral da planilha

COMO USAR:
1. Salva esse arquivo na pasta OfertasTec/
2. Abre PowerShell, vai na pasta e roda:
   python verificar_duplicatas.py

IMPORTANTE: Feche o Excel ANTES de rodar (senão dá Permission denied).
"""

import sys
import re
from pathlib import Path
from collections import defaultdict

try:
    import pandas as pd
except ImportError:
    print("\n❌ ERRO: 'pandas' não instalado.")
    print("Rode no PowerShell: pip install pandas openpyxl")
    sys.exit(1)


ARQUIVO_PLANILHA = "Produtos_ml.xlsx"
ABA_PLANILHA = "Produtos Renomear"

# Categorias oficiais (devem bater EXATAMENTE) — 30 categorias
CATEGORIAS_VALIDAS = {
    "Mouse sem fio",
    "Mouse com fio",
    "Mouse gamer sem fio",
    "Mouse gamer com fio",
    "Mouse vertical (ergonômico)",
    "Teclado com fio",
    "Teclado sem fio",
    "Teclado gamer com fio",
    "Teclado gamer sem fio",
    "Fone sem fio (bluetooth)",
    "Fone com fio",
    "Fone gamer (headset sem fio)",
    "Headset Gamer Com Fio",
    "Caixa de som bluetooth",
    "Caixa de som com fio",
    "Cabo USB-C",
    "Cabo Lightning (iPhone)",
    "Power bank",
    "Suporte p/ celular (carro)",
    "Película protetora",
    "Capinha celular",
    "Mousepad simples",
    "Mousepad gamer (grande)",
    "Controle PC (sem fio)",
    "Controle PC (com fio)",
    "Controle para celular",
    "Adaptador de tomada",
    "Cabo HDMI",
    "Pen drive",
    "HD externo",
}


def normalizar(texto):
    """Remove espaços extras e deixa minúsculo pra comparar"""
    if not texto:
        return ""
    return re.sub(r'\s+', ' ', str(texto).strip().lower())


def main():
    print("\n" + "=" * 65)
    print("  🔍 VERIFICAR DUPLICATAS — OfertasTec")
    print("=" * 65)

    if not Path(ARQUIVO_PLANILHA).exists():
        print(f"\n❌ Planilha '{ARQUIVO_PLANILHA}' não encontrada!")
        print(f"Confirma que ela está na pasta atual.")
        input("\nPressione ENTER pra sair...")
        return

    print(f"\n📊 Lendo '{ARQUIVO_PLANILHA}'...")
    try:
        df = pd.read_excel(ARQUIVO_PLANILHA, sheet_name=ABA_PLANILHA)
    except PermissionError:
        print("\n❌ ERRO: Planilha está ABERTA no Excel!")
        print("Feche o Excel e tente de novo.")
        input("\nPressione ENTER pra sair...")
        return
    except Exception as e:
        print(f"❌ Erro: {e}")
        input("\nPressione ENTER pra sair...")
        return

    print(f"✅ {len(df)} linhas carregadas\n")

    # Coletar dados
    produtos = []
    for idx, row in df.iterrows():
        num = row.get('#')
        cat = row.get('Categoria')
        link = row.get('Link (clique pra abrir)')
        nome = row.get('Nome do Produto')

        if pd.notna(num):
            produtos.append({
                'linha_excel': idx + 2,  # +2 pq idx começa do 0 e tem cabeçalho
                'num': int(num) if pd.notna(num) else None,
                'cat': str(cat).strip() if pd.notna(cat) else '',
                'link': str(link).strip() if pd.notna(link) else '',
                'nome': str(nome).strip() if pd.notna(nome) else '',
            })

    # ═══════════════════════════════════════════════════════
    # VERIFICAÇÃO 1: LINKS DUPLICADOS
    # ═══════════════════════════════════════════════════════
    print("=" * 65)
    print("  1️⃣  LINKS DUPLICADOS")
    print("=" * 65)

    links_map = defaultdict(list)
    for p in produtos:
        if p['link'] and p['link'].startswith('http'):
            links_map[p['link']].append(p)

    duplicatas_link = {k: v for k, v in links_map.items() if len(v) > 1}

    if not duplicatas_link:
        print("  ✅ Nenhum link duplicado!\n")
    else:
        print(f"  ⚠️ {len(duplicatas_link)} link(s) DUPLICADO(s) encontrado(s):\n")
        for link, ocorrencias in duplicatas_link.items():
            print(f"  🔗 {link}")
            for o in ocorrencias:
                print(f"     - Linha Excel {o['linha_excel']} | #{o['num']} | {o['cat']} | {o['nome'][:40] or '(sem nome)'}")
            print()

    # ═══════════════════════════════════════════════════════
    # VERIFICAÇÃO 2: NOMES DUPLICADOS
    # ═══════════════════════════════════════════════════════
    print("=" * 65)
    print("  2️⃣  NOMES DE PRODUTOS DUPLICADOS")
    print("=" * 65)

    nomes_map = defaultdict(list)
    for p in produtos:
        if p['nome']:
            nome_norm = normalizar(p['nome'])
            nomes_map[nome_norm].append(p)

    duplicatas_nome = {k: v for k, v in nomes_map.items() if len(v) > 1}

    if not duplicatas_nome:
        print("  ✅ Nenhum nome duplicado!\n")
    else:
        print(f"  ⚠️ {len(duplicatas_nome)} nome(s) duplicado(s) (pode ser intencional):\n")
        for nome, ocorrencias in duplicatas_nome.items():
            print(f"  📝 \"{ocorrencias[0]['nome']}\"")
            for o in ocorrencias:
                print(f"     - Linha Excel {o['linha_excel']} | #{o['num']} | {o['cat']}")
            print()

    # ═══════════════════════════════════════════════════════
    # VERIFICAÇÃO 3: LINHAS INCOMPLETAS
    # ═══════════════════════════════════════════════════════
    print("=" * 65)
    print("  3️⃣  LINHAS INCOMPLETAS")
    print("=" * 65)

    com_link_sem_nome = [p for p in produtos if p['link'].startswith('http') and not p['nome']]
    com_nome_sem_link = [p for p in produtos if p['nome'] and (not p['link'] or not p['link'].startswith('http'))]

    if not com_link_sem_nome and not com_nome_sem_link:
        print("  ✅ Todas as linhas com link têm nome (e vice-versa)!\n")
    else:
        if com_link_sem_nome:
            print(f"  ⚠️ {len(com_link_sem_nome)} linha(s) com LINK mas SEM NOME:")
            for p in com_link_sem_nome[:15]:
                print(f"     - Linha {p['linha_excel']} | #{p['num']} | {p['cat']} | {p['link']}")
            if len(com_link_sem_nome) > 15:
                print(f"     ... e mais {len(com_link_sem_nome) - 15}")
            print()
        if com_nome_sem_link:
            print(f"  ⚠️ {len(com_nome_sem_link)} linha(s) com NOME mas SEM LINK:")
            for p in com_nome_sem_link[:15]:
                print(f"     - Linha {p['linha_excel']} | #{p['num']} | {p['cat']} | {p['nome'][:50]}")
            if len(com_nome_sem_link) > 15:
                print(f"     ... e mais {len(com_nome_sem_link) - 15}")
            print()

    # ═══════════════════════════════════════════════════════
    # VERIFICAÇÃO 4: CATEGORIAS ESTRANHAS
    # ═══════════════════════════════════════════════════════
    print("=" * 65)
    print("  4️⃣  CATEGORIAS FORA DO PADRÃO")
    print("=" * 65)

    cats_invalidas = defaultdict(list)
    for p in produtos:
        if p['cat'] and p['cat'] not in CATEGORIAS_VALIDAS:
            cats_invalidas[p['cat']].append(p)

    if not cats_invalidas:
        print("  ✅ Todas as categorias estão padronizadas!\n")
    else:
        print(f"  ⚠️ Categorias não reconhecidas:\n")
        for cat, items in cats_invalidas.items():
            print(f"  🏷️  \"{cat}\" (aparece em {len(items)} linha(s))")
            for p in items[:3]:
                print(f"     - Linha {p['linha_excel']} | #{p['num']}")
            if len(items) > 3:
                print(f"     ... e mais {len(items) - 3}")
            print()
        print("  💡 Categorias válidas são:")
        for c in sorted(CATEGORIAS_VALIDAS):
            print(f"     • {c}")
        print()

    # ═══════════════════════════════════════════════════════
    # VERIFICAÇÃO 5: LINKS SUSPEITOS
    # ═══════════════════════════════════════════════════════
    print("=" * 65)
    print("  5️⃣  LINKS SUSPEITOS (não meli.la)")
    print("=" * 65)

    links_suspeitos = [
        p for p in produtos
        if p['link'].startswith('http') and 'meli.la' not in p['link']
    ]

    if not links_suspeitos:
        print("  ✅ Todos os links são do meli.la!\n")
    else:
        print(f"  ⚠️ {len(links_suspeitos)} link(s) NÃO meli.la:\n")
        for p in links_suspeitos[:10]:
            print(f"     - Linha {p['linha_excel']} | #{p['num']} | {p['link'][:60]}")
        print()

    # ═══════════════════════════════════════════════════════
    # RESUMO GERAL
    # ═══════════════════════════════════════════════════════
    print("=" * 65)
    print("  📊 RESUMO GERAL DA PLANILHA")
    print("=" * 65)

    total_linhas = len(produtos)
    com_link = sum(1 for p in produtos if p['link'].startswith('http'))
    com_nome = sum(1 for p in produtos if p['nome'])
    completos = sum(1 for p in produtos if p['link'].startswith('http') and p['nome'])

    print(f"\n  📝 Total de linhas com #:        {total_linhas}")
    print(f"  🔗 Com link válido:              {com_link}")
    print(f"  📛 Com nome preenchido:          {com_nome}")
    print(f"  ✅ Completos (link + nome):      {completos}")
    print(f"  ❌ Faltam preencher:             {com_link - com_nome if com_link > com_nome else 0}")

    print("\n  📂 Produtos por categoria:")
    cats_count = defaultdict(int)
    for p in produtos:
        if p['cat'] and p['link'].startswith('http'):
            cats_count[p['cat']] += 1
    for cat in sorted(cats_count.keys()):
        print(f"     {cat:35} {cats_count[cat]:4}")

    # ═══════════════════════════════════════════════════════
    # RESULTADO FINAL
    # ═══════════════════════════════════════════════════════
    print("\n" + "=" * 65)
    print("  🎯 RESULTADO")
    print("=" * 65)

    total_problemas = (
        len(duplicatas_link) +
        len(com_link_sem_nome) +
        len(com_nome_sem_link) +
        len(cats_invalidas) +
        len(links_suspeitos)
    )

    if total_problemas == 0:
        print("\n  🎉 PLANILHA LIMPA! Nenhum problema crítico encontrado.")
    else:
        print(f"\n  ⚠️ {total_problemas} ponto(s) precisam de atenção (veja acima).")

    print()
    input("Pressione ENTER pra sair...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Saindo...")
    except Exception as e:
        print(f"\n❌ ERRO inesperado: {e}")
        input("\nPressione ENTER pra sair...")
