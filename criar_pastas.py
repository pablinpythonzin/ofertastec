"""
═══════════════════════════════════════════════════════════════════════
  CRIAR PASTAS DE CATEGORIA — OfertasTec
  Cria TODAS as subpastas de categoria de uma vez (definitivo)
═══════════════════════════════════════════════════════════════════════

O QUE FAZ:
  Cria dentro de fotos-webp/ todas as subpastas de categoria com os
  nomes EXATOS que o script de fotos espera. Assim nunca mais dá
  erro de "pasta não existe".

COMO USAR:
  1. Salva na pasta OfertasTec/
  2. Roda: python criar_pastas.py

  Se uma pasta já existe, ele NÃO apaga nada (seguro).
  Se você criou pastas com nome errado antes (cabo-c, head-set-gamer),
  esse script cria as CERTAS ao lado — aí você move as fotos.
═══════════════════════════════════════════════════════════════════════
"""

import sys
from pathlib import Path

PASTA_FOTOS_BRUTAS = "fotos-webp"

# Lista DEFINITIVA: nome-da-pasta : categoria-na-planilha
CATEGORIAS = {
    "mouse-sem-fio":          "Mouse sem fio",
    "mouse-com-fio":          "Mouse com fio",
    "mouse-gamer-sem-fio":    "Mouse gamer sem fio",
    "mouse-gamer-com-fio":    "Mouse gamer com fio",
    "teclado-com-fio":        "Teclado com fio",
    "teclado-sem-fio":        "Teclado sem fio",
    "teclado-gamer-com-fio":  "Teclado gamer com fio",
    "fone-sem-fio":           "Fone sem fio (bluetooth)",
    "fone-com-fio":           "Fone com fio",
    "fone-gamer":             "Fone gamer (headset)",
    "cabo-usbc":              "Cabo USB-C",
    "pen-drive":              "Pen drive",
    "mouse-pad":              "Mouse pad",
}


def main():
    print("\n" + "=" * 60)
    print("  📁 CRIAR PASTAS DE CATEGORIA — OfertasTec")
    print("=" * 60)

    pasta_atual = Path.cwd()
    pasta_brutas = pasta_atual / PASTA_FOTOS_BRUTAS

    # cria a pasta-mãe fotos-webp se não existir
    pasta_brutas.mkdir(exist_ok=True)
    print(f"\n📂 Pasta base: {pasta_brutas}\n")

    criadas = 0
    existentes = 0

    print("  Status das pastas de categoria:\n")
    for slug, nome_planilha in CATEGORIAS.items():
        pasta_cat = pasta_brutas / slug
        if pasta_cat.exists():
            print(f"  ✓ {slug:28} (já existia)")
            existentes += 1
        else:
            pasta_cat.mkdir()
            print(f"  ✅ {slug:28} CRIADA")
            criadas += 1

    print("\n" + "=" * 60)
    print(f"  ✨ {criadas} criadas | {existentes} já existiam")
    print("=" * 60)

    print("\n  📋 Mapa pasta → categoria na planilha:\n")
    for slug, nome in CATEGORIAS.items():
        print(f"     {slug:28} → {nome}")

    print("\n  💡 Agora joga as fotos .webp nas pastas certas,")
    print("     renomeadas com o ID do produto (ex: 10.webp, 11.webp)")
    print()
    input("Pressione ENTER pra sair...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        input("\nPressione ENTER pra sair...")
