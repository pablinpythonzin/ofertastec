"""
═══════════════════════════════════════════════════════════════════════
  ORGANIZAR PRODUTOS — OfertasTec
  Move as fotos das subpastas pra raiz de produtos/ (deixa tudo solto)
═══════════════════════════════════════════════════════════════════════

O QUE FAZ:
  O site lê as fotos SOLTAS em produtos/foto.jpg. Mas algumas fotos
  acabaram dentro de subpastas (produtos/teclado-gamer-com-fio/foto.jpg).
  Este script:
    1. Move TODAS as fotos das subpastas pra raiz de produtos/
    2. Remove as subpastas (agora vazias)
  Assim o site volta a achar as fotos.

COMO USAR:
  1. Salva na pasta OfertasTec/
  2. Roda: python organizar_produtos.py
  3. Confirma com 's'

SEGURANÇA:
  - Mostra o que vai fazer ANTES de mexer (você confirma)
  - Se já existir uma foto com mesmo nome na raiz, ele AVISA e não
    sobrescreve (mantém as duas, renomeando a nova)
  - Só remove subpasta depois de mover tudo de dentro
═══════════════════════════════════════════════════════════════════════
"""

import sys
import shutil
from pathlib import Path

PASTA_DESTINO = "produtos"
EXTS = ['.jpg', '.jpeg', '.png', '.webp', '.gif']


def main():
    print("\n" + "=" * 60)
    print("  📦 ORGANIZAR PRODUTOS — deixar fotos soltas")
    print("=" * 60)

    pasta_atual = Path.cwd()
    pasta = pasta_atual / PASTA_DESTINO

    if not pasta.exists():
        print(f"\n❌ Pasta '{PASTA_DESTINO}/' não encontrada!")
        input("\nENTER pra sair..."); return

    # Acha todas as subpastas dentro de produtos/
    subpastas = [p for p in pasta.iterdir() if p.is_dir()]

    if not subpastas:
        print("\n✅ Não tem subpastas em produtos/. Tudo já está solto!")
        input("\nENTER pra sair..."); return

    # Mapeia o que vai mover
    plano = []      # (origem, destino)
    conflitos = []  # nomes que já existem na raiz
    for sub in subpastas:
        for foto in sub.iterdir():
            if foto.is_file() and foto.suffix.lower() in EXTS:
                destino = pasta / foto.name
                if destino.exists():
                    conflitos.append(foto)
                else:
                    plano.append((foto, destino))

    print(f"\n📂 {len(subpastas)} subpasta(s) encontrada(s)")
    print(f"📸 {len(plano)} foto(s) pra mover pra raiz")
    if conflitos:
        print(f"⚠️ {len(conflitos)} foto(s) com nome JÁ existente na raiz (vão ser renomeadas)")

    print("\n  Vai mover:\n")
    for origem, destino in plano[:20]:
        print(f"    {origem.parent.name}/{origem.name}  →  produtos/{destino.name}")
    if len(plano) > 20:
        print(f"    ... e mais {len(plano) - 20} foto(s)")

    if conflitos:
        print("\n  ⚠️ Conflitos (nome repetido, vão virar nome-2.jpg):\n")
        for f in conflitos[:10]:
            print(f"    {f.parent.name}/{f.name}")

    print()
    if input("  ✅ Confirma mover tudo pra raiz? (s/n): ").strip().lower() != 's':
        print("  Cancelado."); input("\n  ENTER..."); return

    print("\n  ⚙️  Movendo...\n")
    movidas = 0

    # Move os sem conflito
    for origem, destino in plano:
        shutil.move(str(origem), str(destino))
        print(f"  ✅ {destino.name}")
        movidas += 1

    # Move os com conflito, renomeando
    for f in conflitos:
        base = f.stem
        ext = f.suffix
        n = 2
        novo_destino = pasta / f"{base}-{n}{ext}"
        while novo_destino.exists():
            n += 1
            novo_destino = pasta / f"{base}-{n}{ext}"
        shutil.move(str(f), str(novo_destino))
        print(f"  ✅ {novo_destino.name} (renomeado p/ não sobrescrever)")
        movidas += 1

    # Remove as subpastas (agora vazias)
    print("\n  🗑️  Removendo subpastas vazias...\n")
    removidas = 0
    for sub in subpastas:
        try:
            # só remove se estiver vazia
            restante = [x for x in sub.iterdir()]
            if not restante:
                sub.rmdir()
                print(f"  ✅ removida: {sub.name}/")
                removidas += 1
            else:
                print(f"  ⚠️ {sub.name}/ ainda tem arquivos (não removida): {[x.name for x in restante]}")
        except Exception as e:
            print(f"  ⚠️ erro ao remover {sub.name}/: {e}")

    print("\n" + "=" * 60)
    print(f"  ✨ {movidas} foto(s) movidas | {removidas} subpasta(s) removidas")
    print("=" * 60)
    print("\n  ✅ Agora as fotos estão soltas em produtos/ — o site vai achar!")
    print("\n  📝 Próximo: git add . && git commit && git push")
    print()
    input("ENTER pra sair...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Saindo...")
    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        input("\nENTER pra sair...")
