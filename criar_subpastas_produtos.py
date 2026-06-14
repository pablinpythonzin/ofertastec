"""
═══════════════════════════════════════════════════════════════════════
  CRIAR SUBPASTAS EM PRODUTOS — OfertasTec
  Cria dentro de produtos/ as mesmas subpastas que existem em fotos-webp/
═══════════════════════════════════════════════════════════════════════

O QUE FAZ:
  Olha quais subpastas de categoria existem em fotos-webp/ e cria as
  mesmas dentro de produtos/. Assim as duas pastas ficam com a mesma
  estrutura de categorias.

  Se você não tiver a pasta fotos-webp/ ainda, ele usa a lista fixa
  das 30 categorias oficiais.

COMO USAR:
  1. Salva na pasta OfertasTec/
  2. Roda: python criar_subpastas_produtos.py

  Seguro: se a subpasta já existe, ele NÃO apaga nada.

⚠️ LEMBRETE: o site lê as fotos SOLTAS em produtos/ (produtos/foto.jpg).
  Estas subpastas ficam criadas pra organização, mas as fotos que o
  site usa continuam soltas na raiz de produtos/.
═══════════════════════════════════════════════════════════════════════
"""

import sys
from pathlib import Path

PASTA_FOTOS_BRUTAS = "fotos-webp"
PASTA_DESTINO = "produtos"

# Lista fixa das 30 categorias (usada se fotos-webp/ não existir)
CATEGORIAS_PADRAO = [
    "mouse-sem-fio", "mouse-com-fio", "mouse-gamer-sem-fio", "mouse-gamer-com-fio",
    "mouse-vertical", "teclado-com-fio", "teclado-sem-fio", "teclado-gamer-com-fio",
    "teclado-gamer-sem-fio", "fone-sem-fio", "fone-com-fio", "fone-gamer-sem-fio",
    "fone-gamer-com-fio", "caixa-som-bluetooth", "caixa-som-com-fio", "cabo-usbc",
    "cabo-lightning", "power-bank", "suporte-celular-carro", "pelicula-protetora",
    "capinha-celular", "mousepad-simples", "mousepad-gamer", "controle-pc-sem-fio",
    "controle-pc-com-fio", "controle-celular", "adaptador-tomada", "cabo-hdmi",
    "pen-drive", "hd-externo",
]


def main():
    print("\n" + "=" * 60)
    print("  📁 CRIAR SUBPASTAS EM PRODUTOS — OfertasTec")
    print("=" * 60)

    pasta_atual = Path.cwd()
    pasta_brutas = pasta_atual / PASTA_FOTOS_BRUTAS
    pasta_destino = pasta_atual / PASTA_DESTINO

    # Garante que produtos/ existe
    pasta_destino.mkdir(exist_ok=True)
    print(f"\n📂 Pasta destino: {pasta_destino}")

    # Descobre quais categorias usar
    if pasta_brutas.exists():
        categorias = sorted([
            p.name for p in pasta_brutas.iterdir() if p.is_dir()
        ])
        if categorias:
            print(f"✅ Lendo {len(categorias)} categorias de fotos-webp/\n")
        else:
            print("⚠️ fotos-webp/ está vazia, usando lista padrão das 30\n")
            categorias = CATEGORIAS_PADRAO
    else:
        print("⚠️ fotos-webp/ não encontrada, usando lista padrão das 30\n")
        categorias = CATEGORIAS_PADRAO

    # Cria as subpastas em produtos/
    criadas = 0
    existentes = 0
    print("  Status das subpastas em produtos/:\n")
    for slug in categorias:
        sub = pasta_destino / slug
        if sub.exists():
            print(f"  ✓ {slug:24} (já existia)")
            existentes += 1
        else:
            sub.mkdir()
            print(f"  ✅ {slug:24} CRIADA")
            criadas += 1

    print("\n" + "=" * 60)
    print(f"  ✨ {criadas} criadas | {existentes} já existiam | {len(categorias)} no total")
    print("=" * 60)

    print("\n  ⚠️ Lembrete: o site lê fotos SOLTAS em produtos/foto.jpg")
    print("     Estas subpastas são só pra organização.")
    print()
    input("Pressione ENTER pra sair...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        input("\nPressione ENTER pra sair...")
