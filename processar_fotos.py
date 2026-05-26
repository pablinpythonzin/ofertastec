"""
═══════════════════════════════════════════════════════════════
  PROCESSAR FOTOS — OfertasTec
  Renomeia e converte fotos .webp/.jpg pra usar no site
═══════════════════════════════════════════════════════════════

COMO USAR:
1. Coloca esse arquivo na pasta OfertasTec/
2. Cria pasta 'fotos-webp/' e joga TODAS as fotos baixadas lá
3. Roda no PowerShell:
   pip install pandas openpyxl pillow
   python processar_fotos.py

O QUE FAZ:
- Lê fotos da pasta 'fotos-webp/'
- Pra cada uma, ABRE pra você ver
- Você digita o número do produto (da planilha)
- Script renomeia, converte e salva em 'produtos/'

CONTROLES NO PROMPT:
- Digite o número do produto (1-140)
- 'p' = pular essa foto (não usar)
- 's' = sair do script
"""

import os
import sys
import unicodedata
import re
import shutil
from pathlib import Path

# ───────────────────────────────────────────────────────────
# Verificar dependências
# ───────────────────────────────────────────────────────────
try:
    import pandas as pd
except ImportError:
    print("\n❌ ERRO: 'pandas' não instalado.")
    print("Rode no PowerShell: pip install pandas openpyxl pillow")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("\n❌ ERRO: 'Pillow' não instalado.")
    print("Rode no PowerShell: pip install pillow")
    sys.exit(1)


# ───────────────────────────────────────────────────────────
# Configurações
# ───────────────────────────────────────────────────────────
PASTA_FOTOS_BRUTAS = "fotos-webp"
PASTA_DESTINO = "produtos"
ARQUIVO_PLANILHA = "Produtos_ml.xlsx"
ABA_PLANILHA = "Produtos Renomear"


# ───────────────────────────────────────────────────────────
# Funções utilitárias
# ───────────────────────────────────────────────────────────
def remover_acentos(texto):
    """Remove acentos: 'ergonômico' → 'ergonomico'"""
    if not texto:
        return ""
    nfkd = unicodedata.normalize('NFKD', str(texto))
    return ''.join([c for c in nfkd if not unicodedata.combining(c)])


def gerar_nome_arquivo(nome_produto):
    """Aplica padrão web no nome: minúsculo + hifens + sem acento"""
    nome = remover_acentos(nome_produto)
    nome = nome.lower()
    nome = re.sub(r'[^a-z0-9]+', '-', nome)  # tudo que não for letra/número vira hífen
    nome = re.sub(r'-+', '-', nome)  # múltiplos hífens viram um só
    nome = nome.strip('-')  # tira hífens nas pontas
    return nome


def abrir_imagem(caminho):
    """Abre a imagem no visualizador padrão do Windows"""
    try:
        os.startfile(caminho)  # Windows
    except AttributeError:
        # se não for Windows, tenta outras
        import subprocess
        subprocess.run(['xdg-open', caminho], check=False)


def converter_e_salvar(caminho_origem, caminho_destino):
    """Converte qualquer formato pra .jpg de qualidade boa"""
    try:
        img = Image.open(caminho_origem)
        # se tiver transparência, converte pra RGB com fundo branco
        if img.mode in ('RGBA', 'LA', 'P'):
            fundo = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            fundo.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
            img = fundo
        elif img.mode != 'RGB':
            img = img.convert('RGB')

        # salva com qualidade boa e otimizado
        img.save(caminho_destino, 'JPEG', quality=85, optimize=True)
        tamanho_kb = os.path.getsize(caminho_destino) / 1024
        return True, f"{tamanho_kb:.1f} KB"
    except Exception as e:
        return False, str(e)


# ───────────────────────────────────────────────────────────
# Script principal
# ───────────────────────────────────────────────────────────
def main():
    print("\n" + "=" * 60)
    print("  PROCESSAR FOTOS — OfertasTec")
    print("=" * 60)

    # 1. Verificar pastas e planilha
    pasta_atual = Path.cwd()
    pasta_brutas = pasta_atual / PASTA_FOTOS_BRUTAS
    pasta_destino = pasta_atual / PASTA_DESTINO

    if not pasta_brutas.exists():
        print(f"\n❌ Pasta '{PASTA_FOTOS_BRUTAS}/' não existe!")
        print(f"Crie a pasta e coloque as fotos dentro: {pasta_brutas}")
        input("\nPressione ENTER pra sair...")
        return

    if not Path(ARQUIVO_PLANILHA).exists():
        print(f"\n❌ Planilha '{ARQUIVO_PLANILHA}' não encontrada!")
        print(f"Confirme que ela está em: {pasta_atual}")
        input("\nPressione ENTER pra sair...")
        return

    # criar pasta destino se não existir
    pasta_destino.mkdir(exist_ok=True)

    # 2. Carregar planilha
    print(f"\n📊 Lendo planilha '{ARQUIVO_PLANILHA}'...")
    try:
        df = pd.read_excel(ARQUIVO_PLANILHA, sheet_name=ABA_PLANILHA)
    except Exception as e:
        print(f"❌ Erro ao ler planilha: {e}")
        input("\nPressione ENTER pra sair...")
        return

    # dicionário: número → nome do produto
    produtos = {}
    for _, row in df.iterrows():
        num = row.get('#')
        nome = row.get('Nome do Produto')
        if pd.notna(num) and pd.notna(nome) and str(nome).strip():
            produtos[int(num)] = str(nome).strip()

    print(f"✅ {len(produtos)} produtos carregados\n")

    # 3. Listar fotos da pasta brutas
    extensoes = ['.webp', '.jpg', '.jpeg', '.png']
    fotos = sorted([
        f for f in pasta_brutas.iterdir()
        if f.is_file() and f.suffix.lower() in extensoes
    ])

    if not fotos:
        print(f"❌ Nenhuma foto encontrada em '{PASTA_FOTOS_BRUTAS}/'")
        print(f"Formatos aceitos: {', '.join(extensoes)}")
        input("\nPressione ENTER pra sair...")
        return

    print(f"📸 {len(fotos)} fotos encontradas em '{PASTA_FOTOS_BRUTAS}/'\n")

    # 4. Processar foto por foto
    processadas = 0
    puladas = 0
    erros = 0

    for i, foto in enumerate(fotos, 1):
        print("\n" + "─" * 60)
        print(f"  📷 FOTO {i}/{len(fotos)}: {foto.name}")
        print("─" * 60)
        print("  [Abrindo no visualizador... olhe a foto]")

        abrir_imagem(str(foto))

        # pergunta o número do produto
        while True:
            resposta = input(
                "\n  → Digite o # do produto (1-140), 'p' pra pular, 's' pra sair: "
            ).strip().lower()

            if resposta == 's':
                print("\n👋 Saindo...")
                _imprimir_resumo(processadas, puladas, erros, len(fotos))
                return

            if resposta == 'p':
                print("  ⏭️  Foto pulada.")
                puladas += 1
                break

            try:
                num = int(resposta)
                if num not in produtos:
                    print(f"  ⚠️ Produto #{num} não encontrado na planilha. Tenta de novo.")
                    continue

                nome_produto = produtos[num]
                nome_arquivo = gerar_nome_arquivo(nome_produto) + '.jpg'
                caminho_final = pasta_destino / nome_arquivo

                # confirma com o usuário
                print(f"\n  Produto #{num}: {nome_produto}")
                print(f"  Arquivo vai virar: {nome_arquivo}")

                if caminho_final.exists():
                    sobrescrever = input(
                        "  ⚠️ Já existe um arquivo com esse nome. Sobrescrever? (s/n): "
                    ).strip().lower()
                    if sobrescrever != 's':
                        print("  ⏭️ Pulando esta foto.")
                        puladas += 1
                        break

                # converte e salva
                print("  ⚙️  Convertendo e otimizando...")
                ok, info = converter_e_salvar(str(foto), str(caminho_final))

                if ok:
                    print(f"  ✅ Salvo em produtos/{nome_arquivo} ({info})")
                    processadas += 1
                else:
                    print(f"  ❌ Erro ao processar: {info}")
                    erros += 1

                break

            except ValueError:
                print("  ⚠️ Digite um número válido (ou 'p' pra pular, 's' pra sair).")

    # 5. Resumo final
    _imprimir_resumo(processadas, puladas, erros, len(fotos))


def _imprimir_resumo(processadas, puladas, erros, total):
    print("\n" + "=" * 60)
    print("  ✨ RESUMO")
    print("=" * 60)
    print(f"  ✅ Processadas: {processadas}/{total}")
    print(f"  ⏭️  Puladas:    {puladas}")
    print(f"  ❌ Erros:       {erros}")
    print("=" * 60)
    print(f"\n📁 Fotos prontas em: produtos/")
    print(f"📝 Próximo passo: mande o Claude atualizar o index.html\n")

    input("Pressione ENTER pra sair...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Saindo (Ctrl+C)...")
    except Exception as e:
        print(f"\n❌ ERRO inesperado: {e}")
        input("\nPressione ENTER pra sair...")
