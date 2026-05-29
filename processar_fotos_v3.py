"""
═══════════════════════════════════════════════════════════════════════
  PROCESSAR FOTOS v3 — OfertasTec
  Organiza fotos por SUBPASTAS de categoria
═══════════════════════════════════════════════════════════════════════

COMO USAR:
1. Cria subpastas em fotos-webp/ com nomes:
   - mouse-sem-fio
   - mouse-com-fio
   - mouse-gamer-sem-fio
   - mouse-gamer-com-fio
   - teclado-com-fio
   - teclado-sem-fio
   - teclado-gamer-com-fio
   - fone-sem-fio
   - fone-com-fio
   - fone-gamer
   - cabo-usbc
   - pen-drive

2. Joga as fotos brutas (.webp) nas pastas certas
3. Roda o script:
     python processar_fotos_v3.py

4. Script:
   - Mostra menu das categorias
   - Você escolhe uma
   - Lista produtos daquela categoria (✅ já tem foto, ⏳ falta)
   - Abre cada foto e você digita o # do produto
   - Renomeia, converte .webp→.jpg e salva em produtos/

ATALHOS DURANTE PROCESSAMENTO:
  p  = pular foto
  s  = sair (voltar ao menu)
  v  = ver foto de novo
  l  = listar produtos novamente

═══════════════════════════════════════════════════════════════════════
"""

import os
import sys
import unicodedata
import re
from pathlib import Path

# Dependências
try:
    import pandas as pd
except ImportError:
    print("\n❌ ERRO: 'pandas' não instalado.")
    print("Rode: pip install pandas openpyxl pillow")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("\n❌ ERRO: 'Pillow' não instalado.")
    print("Rode: pip install pillow")
    sys.exit(1)


# ═══════════════════════════════════════════════════════════════════════
# CONFIGURAÇÕES
# ═══════════════════════════════════════════════════════════════════════

PASTA_FOTOS_BRUTAS = "fotos-webp"
PASTA_DESTINO = "produtos"
ARQUIVO_PLANILHA = "Produtos_ml.xlsx"
ABA_PLANILHA = "Produtos Renomear"

# Mapeamento: nome-pasta → nome-na-planilha
CATEGORIAS = {
    "mouse-sem-fio": "Mouse sem fio",
    "mouse-com-fio": "Mouse com fio",
    "mouse-gamer-sem-fio": "Mouse gamer sem fio",
    "mouse-gamer-com-fio": "Mouse gamer com fio",
    "teclado-com-fio": "Teclado com fio",
    "teclado-sem-fio": "Teclado sem fio",
    "teclado-gamer-com-fio": "Teclado gamer com fio",
    "fone-sem-fio": "Fone sem fio (bluetooth)",
    "fone-com-fio": "Fone com fio",
    "fone-gamer": "Fone gamer (headset)",
    "cabo-usbc": "Cabo USB-C",
    "pen-drive": "Pen drive",
}


# ═══════════════════════════════════════════════════════════════════════
# FUNÇÕES UTILITÁRIAS
# ═══════════════════════════════════════════════════════════════════════

def remover_acentos(texto):
    if not texto:
        return ""
    nfkd = unicodedata.normalize('NFKD', str(texto))
    return ''.join([c for c in nfkd if not unicodedata.combining(c)])


def gerar_nome_arquivo(nome_produto):
    nome = remover_acentos(nome_produto)
    nome = nome.lower()
    nome = re.sub(r'[^a-z0-9]+', '-', nome)
    nome = re.sub(r'-+', '-', nome)
    nome = nome.strip('-')
    return nome


def abrir_imagem(caminho):
    try:
        os.startfile(caminho)
    except AttributeError:
        import subprocess
        subprocess.run(['xdg-open', caminho], check=False)


def converter_e_salvar(caminho_origem, caminho_destino):
    try:
        img = Image.open(caminho_origem)
        if img.mode in ('RGBA', 'LA', 'P'):
            fundo = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            fundo.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
            img = fundo
        elif img.mode != 'RGB':
            img = img.convert('RGB')
        img.save(caminho_destino, 'JPEG', quality=85, optimize=True)
        tamanho_kb = os.path.getsize(caminho_destino) / 1024
        return True, f"{tamanho_kb:.1f} KB"
    except Exception as e:
        return False, str(e)


def carregar_planilha():
    if not Path(ARQUIVO_PLANILHA).exists():
        print(f"\n❌ Planilha '{ARQUIVO_PLANILHA}' não encontrada!")
        return None

    try:
        df = pd.read_excel(ARQUIVO_PLANILHA, sheet_name=ABA_PLANILHA)
    except PermissionError:
        print("\n❌ ERRO: Planilha aberta no Excel! Feche e tente de novo.")
        return None
    except Exception as e:
        print(f"❌ Erro: {e}")
        return None

    produtos_por_cat = {}
    for _, row in df.iterrows():
        num = row.get('#')
        cat = row.get('Categoria')
        nome = row.get('Nome do Produto')
        if pd.notna(num) and pd.notna(cat) and pd.notna(nome) and str(nome).strip():
            cat_str = str(cat).strip()
            produtos_por_cat.setdefault(cat_str, []).append({
                'num': int(num),
                'nome': str(nome).strip(),
            })
    return produtos_por_cat


def status_categoria(pasta_categoria, produtos_categoria, pasta_destino):
    """Retorna (n_fotos_brutas, n_produtos_total, n_produtos_com_foto)"""
    extensoes = ['.webp', '.jpg', '.jpeg', '.png']
    fotos_brutas = [
        f for f in pasta_categoria.iterdir()
        if f.is_file() and f.suffix.lower() in extensoes
    ] if pasta_categoria.exists() else []

    n_com_foto = 0
    for p in produtos_categoria:
        nome_jpg = gerar_nome_arquivo(p['nome']) + '.jpg'
        if (pasta_destino / nome_jpg).exists():
            n_com_foto += 1

    return len(fotos_brutas), len(produtos_categoria), n_com_foto


# ═══════════════════════════════════════════════════════════════════════
# MENU PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════

def mostrar_menu(produtos_por_cat, pasta_brutas, pasta_destino):
    print("\n" + "=" * 65)
    print("  📁 CATEGORIAS DISPONÍVEIS PARA PROCESSAR")
    print("=" * 65)
    print()
    print(f"  {'#':>3} | {'Categoria':30} | {'Fotos':>7} | {'Progresso':>20}")
    print("  " + "─" * 75)

    opcoes = []
    for i, (slug, nome_planilha) in enumerate(CATEGORIAS.items(), 1):
        pasta_cat = pasta_brutas / slug
        produtos_cat = produtos_por_cat.get(nome_planilha, [])
        n_fotos, n_total, n_feitos = status_categoria(pasta_cat, produtos_cat, pasta_destino)

        if n_total == 0:
            progresso = "(sem produtos)"
        else:
            pct = (n_feitos / n_total) * 100
            progresso = f"{n_feitos}/{n_total} ({pct:.0f}%)"

        if not pasta_cat.exists():
            status_pasta = "❌ pasta n/existe"
        elif n_fotos == 0:
            status_pasta = "vazia"
        else:
            status_pasta = f"{n_fotos}"

        print(f"  [{i:>2}] | {slug:30} | {status_pasta:>7} | {progresso:>20}")
        opcoes.append((slug, nome_planilha, produtos_cat))

    print()
    print(f"  [0] 🚪 SAIR")
    print()

    while True:
        resposta = input("  Escolha categoria [0 pra sair]: ").strip()
        try:
            num = int(resposta)
            if num == 0:
                return None
            if 1 <= num <= len(opcoes):
                return opcoes[num - 1]
            print(f"  ⚠️ Digite um número de 0 a {len(opcoes)}.")
        except ValueError:
            print("  ⚠️ Digite um número.")


# ═══════════════════════════════════════════════════════════════════════
# PROCESSAR UMA CATEGORIA
# ═══════════════════════════════════════════════════════════════════════

def listar_produtos(produtos, pasta_destino):
    print("\n  📋 Produtos dessa categoria na planilha:")
    print()
    for p in produtos:
        nome_jpg = gerar_nome_arquivo(p['nome']) + '.jpg'
        status = "✅" if (pasta_destino / nome_jpg).exists() else "⏳"
        nome_curto = p['nome'][:50]
        print(f"     {status} #{p['num']:>3} - {nome_curto}")
    print()


def processar_categoria(slug, nome_planilha, produtos, pasta_brutas, pasta_destino):
    pasta_cat = pasta_brutas / slug

    if not pasta_cat.exists():
        print(f"\n  ❌ Pasta '{pasta_brutas}/{slug}' não existe!")
        print(f"  Crie a pasta e coloque as fotos dentro.")
        input("\n  Pressione ENTER...")
        return

    extensoes = ['.webp', '.jpg', '.jpeg', '.png']
    fotos = sorted([
        f for f in pasta_cat.iterdir()
        if f.is_file() and f.suffix.lower() in extensoes
    ])

    if not fotos:
        print(f"\n  ❌ Nenhuma foto em '{pasta_brutas}/{slug}/'")
        input("\n  Pressione ENTER...")
        return

    print("\n" + "=" * 65)
    print(f"  📁 PROCESSANDO: {slug}")
    print("=" * 65)
    print(f"\n  📂 Categoria na planilha: {nome_planilha}")
    print(f"  📸 Fotos pra processar: {len(fotos)}")
    print(f"  🎯 Produtos da categoria: {len(produtos)}")

    if not produtos:
        print(f"\n  ❌ Nenhum produto encontrado pra '{nome_planilha}' na planilha.")
        print("  Confira se a categoria está escrita certo na coluna C.")
        input("\n  Pressione ENTER...")
        return

    listar_produtos(produtos, pasta_destino)

    # Set de #s válidos pra checagem
    nums_validos = {p['num']: p for p in produtos}

    processadas = 0
    puladas = 0
    erros = 0

    for i, foto in enumerate(fotos, 1):
        print("\n" + "─" * 65)
        print(f"  📷 FOTO {i}/{len(fotos)}: {foto.name}")
        print("─" * 65)
        print("  [Abrindo no visualizador...]")

        abrir_imagem(str(foto))

        while True:
            resposta = input(
                "\n  → # do produto | (p)ular | (s)air | (v)er foto | (l)istar: "
            ).strip().lower()

            if resposta == 's':
                print("\n  👋 Voltando ao menu...")
                _resumo(processadas, puladas, erros, len(fotos))
                return

            if resposta == 'p':
                print("  ⏭️  Pulada.")
                puladas += 1
                break

            if resposta == 'v':
                abrir_imagem(str(foto))
                continue

            if resposta == 'l':
                listar_produtos(produtos, pasta_destino)
                continue

            try:
                num = int(resposta)
                if num not in nums_validos:
                    print(f"  ⚠️ Produto #{num} não pertence a essa categoria (ou tá vermelho).")
                    print(f"  Use 'l' pra ver os disponíveis.")
                    continue

                produto = nums_validos[num]
                nome_arquivo = gerar_nome_arquivo(produto['nome']) + '.jpg'
                caminho_final = pasta_destino / nome_arquivo

                print(f"\n  Produto #{num}: {produto['nome']}")
                print(f"  Arquivo: {nome_arquivo}")

                if caminho_final.exists():
                    sobrescrever = input(
                        "  ⚠️ Já existe arquivo com esse nome. Sobrescrever? (s/n): "
                    ).strip().lower()
                    if sobrescrever != 's':
                        print("  ⏭️ Pulando.")
                        puladas += 1
                        break

                print("  ⚙️  Convertendo...")
                ok, info = converter_e_salvar(str(foto), str(caminho_final))

                if ok:
                    print(f"  ✅ Salvo ({info})")
                    processadas += 1
                else:
                    print(f"  ❌ Erro: {info}")
                    erros += 1
                break

            except ValueError:
                print("  ⚠️ Use um número, 'p', 's', 'v' ou 'l'.")

    _resumo(processadas, puladas, erros, len(fotos))
    input("\n  Pressione ENTER pra voltar ao menu...")


def _resumo(processadas, puladas, erros, total):
    print("\n" + "─" * 65)
    print("  ✨ RESUMO DA CATEGORIA")
    print("─" * 65)
    print(f"  ✅ Processadas: {processadas}/{total}")
    print(f"  ⏭️  Puladas:    {puladas}")
    print(f"  ❌ Erros:       {erros}")


# ═══════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════

def main():
    print("\n" + "=" * 65)
    print("  🚀 PROCESSAR FOTOS v3 — OfertasTec")
    print("=" * 65)

    pasta_atual = Path.cwd()
    pasta_brutas = pasta_atual / PASTA_FOTOS_BRUTAS
    pasta_destino = pasta_atual / PASTA_DESTINO

    if not pasta_brutas.exists():
        print(f"\n❌ Pasta '{PASTA_FOTOS_BRUTAS}/' não existe!")
        input("\nPressione ENTER...")
        return

    pasta_destino.mkdir(exist_ok=True)

    print(f"\n📊 Carregando planilha...")
    produtos_por_cat = carregar_planilha()
    if produtos_por_cat is None:
        input("\nPressione ENTER...")
        return

    total = sum(len(v) for v in produtos_por_cat.values())
    print(f"✅ {total} produtos carregados em {len(produtos_por_cat)} categorias")

    # Loop principal
    while True:
        escolha = mostrar_menu(produtos_por_cat, pasta_brutas, pasta_destino)
        if escolha is None:
            print("\n  👋 Saindo. Até a próxima!\n")
            break
        slug, nome_planilha, produtos = escolha
        processar_categoria(slug, nome_planilha, produtos, pasta_brutas, pasta_destino)

    print("\n📁 Fotos prontas em: produtos/")
    print("📝 Próximo passo: peça ao Claude pra atualizar o index.html\n")
    input("Pressione ENTER pra sair...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Saindo (Ctrl+C)...")
    except Exception as e:
        print(f"\n❌ ERRO inesperado: {e}")
        input("\nPressione ENTER pra sair...")
