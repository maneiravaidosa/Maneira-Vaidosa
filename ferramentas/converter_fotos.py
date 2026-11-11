"""
Converte as fotos da pasta "fotos originais" em PNG com fundo transparente
para o site (pasta png/).

  fotos originais/20_castanho.jpg  ->  png/20/20 castanho.png           (foto grande, página do modelo)
                                   ->  png/20/mini/20 castanho.webp     (miniatura leve, grelha)

Uso (na pasta do projeto):
    pip install "rembg[cpu]" pillow
    python ferramentas/converter_fotos.py           # só converte fotos novas ou alteradas
    python ferramentas/converter_fotos.py --todas   # volta a converter tudo

Da primeira vez, o rembg descarrega automaticamente o modelo de recorte (~180 MB).
"""
import os
import re
import sys
from pathlib import Path

from PIL import Image, ImageOps

RAIZ = Path(__file__).resolve().parent.parent
ORIGEM = RAIZ / "fotos originais"
DESTINO = RAIZ / "png"
LARGURA_MAX = 1000   # px do PNG final (lado maior)
LARGURA_MINI = 520   # px da miniatura usada na grelha
ENTRADA_MAX = 2400   # px da foto enviada ao recorte
MARGEM = 0.04        # margem transparente à volta do sapato
EXTENSOES = {".jpg", ".jpeg", ".png", ".webp", ".heic"}


def nome_destino(ficheiro: Path):
    """'20_Castanho.jpg' ou '20 castanho.jpg' -> ('20', 'castanho')"""
    m = re.match(r"^\s*([^_\s]+)[_\s]+(.+?)\s*$", ficheiro.stem)
    if not m:
        return None
    return m.group(1), m.group(2).lower()


def recortar(img: Image.Image) -> Image.Image:
    alpha = img.getchannel("A").point(lambda a: 255 if a > 12 else 0)
    caixa = alpha.getbbox()
    if not caixa:
        return img
    img = img.crop(caixa)
    w, h = img.size
    m = int(max(w, h) * MARGEM)
    tela = Image.new("RGBA", (w + 2 * m, h + 2 * m), (0, 0, 0, 0))
    tela.paste(img, (m, m))
    return tela


def main():
    todas = "--todas" in sys.argv
    os.environ.setdefault("OMP_NUM_THREADS", "4")
    from rembg import new_session, remove

    ficheiros = sorted(f for f in ORIGEM.iterdir() if f.suffix.lower() in EXTENSOES)
    sessao = None
    feitos = ignorados = 0
    for f in ficheiros:
        alvo = nome_destino(f)
        if not alvo:
            print(f"  ? nome não reconhecido: {f.name}")
            continue
        ref, cor = alvo
        saida = DESTINO / ref / f"{ref} {cor}.png"
        if not todas and saida.exists() and saida.stat().st_mtime >= f.stat().st_mtime:
            ignorados += 1
            continue
        if sessao is None:
            sessao = new_session("isnet-general-use")
        img = ImageOps.exif_transpose(Image.open(f)).convert("RGB")
        img.thumbnail((ENTRADA_MAX, ENTRADA_MAX), Image.LANCZOS)
        img = recortar(remove(img, session=sessao))
        img.thumbnail((LARGURA_MAX, LARGURA_MAX), Image.LANCZOS)
        saida.parent.mkdir(parents=True, exist_ok=True)
        img.save(saida, optimize=True)
        mini = DESTINO / ref / "mini" / f"{ref} {cor}.webp"
        mini.parent.mkdir(parents=True, exist_ok=True)
        pequena = img.copy()
        pequena.thumbnail((LARGURA_MINI, LARGURA_MINI), Image.LANCZOS)
        pequena.save(mini, quality=86, method=6)
        feitos += 1
        print(f"  ✓ {f.name} -> {saida.relative_to(RAIZ)}")
    print(f"\nConvertidas: {feitos} · já estavam feitas: {ignorados}")


if __name__ == "__main__":
    main()
