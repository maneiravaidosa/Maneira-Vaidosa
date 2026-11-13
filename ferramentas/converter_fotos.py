"""
Converte as fotos da pasta "fotos originais" em PNG com fundo transparente
para o site (pasta png/).

  fotos originais/20_castanho.jpg  ->  png/20/20 castanho.png           (foto grande, página do modelo)
                                   ->  png/20/mini/20 castanho.webp     (miniatura leve, grelha)

Uso (na pasta do projeto):
    pip install "rembg[cpu]" pillow
    python ferramentas/converter_fotos.py           # só converte fotos novas ou alteradas
    python ferramentas/converter_fotos.py --todas   # volta a converter tudo
    python ferramentas/converter_fotos.py 20_castanho.jpg   # volta a converter só essa foto
    python ferramentas/converter_fotos.py --ref 1128         # volta a converter todas as fotos dessa referência

Da primeira vez, o rembg descarrega automaticamente o modelo de recorte (~180 MB).
"""
import os
import re
import sys
from pathlib import Path

from PIL import Image, ImageChops, ImageFilter, ImageOps

RAIZ = Path(__file__).resolve().parent.parent
ORIGEM = RAIZ / "fotos originais"
DESTINO = RAIZ / "png"
LARGURA_MAX = 1000   # px do PNG final (lado maior)
LARGURA_MINI = 520   # px da miniatura usada na grelha
ENTRADA_MAX = 2400   # px da foto enviada ao recorte
MARGEM = 0.04        # margem transparente à volta do sapato
EXTENSOES = {".jpg", ".jpeg", ".png", ".webp", ".heic"}

# Modelo de recorte usado por omissão e exceções por referência.
# O "isnet-general-use" funciona bem na maioria das fotos, mas nos chinelos de sola de cortiça
# deixa a sombra do chão como um "nevoeiro" por baixo da sola; para esses usa-se o "u2net"
# e recupera-se a sola preta pela cor (o recorte confunde-a com a sombra e corta-a).
MODELO = "isnet-general-use"
MODELO_POR_REF = {
    "1128": "u2net",
    "2072": "u2net",
    "2703": "u2net",
    "3521": "u2net",
    "3521C": "u2net",
    # partes brancas encostadas à parede branca (ex.: ponta da correia): junta os dois recortes
    "2060": "isnet-general-use+u2net",
}


def nome_destino(ficheiro: Path):
    """'20_Castanho.jpg' ou '20 castanho.jpg' -> ('20', 'castanho')"""
    m = re.match(r"^\s*([^_\s]+)[_\s]+(.+?)\s*$", ficheiro.stem)
    if not m:
        return None
    return m.group(1), m.group(2).lower()


def recuperar_sola(original: Image.Image, cortada: Image.Image, escuro_max=70, passos=60) -> Image.Image:
    """Volta a juntar a sola preta que o recorte cortou.

    A partir do sapato recortado, "cresce" pelos píxeis muito escuros da foto original que lhe estão ligados
    (a sola de borracha); a sombra do chão é mais clara e trava o crescimento, por isso não entra.
    No fim as arestas ficam nítidas: o "quase sapato" fica opaco e o "quase fundo" (nevoeiro) desaparece."""
    a = cortada.getchannel("A")
    escuro = original.convert("L").point(lambda v: 255 if v < escuro_max else 0)
    m = a.point(lambda v: 255 if v > 128 else 0)
    for _ in range(passos):
        novo = ImageChops.lighter(m, ImageChops.multiply(m.filter(ImageFilter.MaxFilter(5)), escuro))
        if ImageChops.difference(novo, m).getbbox() is None:
            break
        m = novo
    m = ImageChops.lighter(a, m).filter(ImageFilter.MedianFilter(5))
    m = m.point(lambda v: 0 if v <= 70 else 255 if v >= 190 else int((v - 70) * 255 / 120))
    cortada.putalpha(m.filter(ImageFilter.GaussianBlur(0.8)))
    return cortada


def juntar_recortes(original: Image.Image, recortes) -> Image.Image:
    """Junta vários recortes: fica sapato o que qualquer um deles achar que é sapato.

    Útil quando uma parte branca do sapato está à frente da parede branca e um dos recortes a "rói"."""
    a = recortes[0].getchannel("A")
    for r in recortes[1:]:
        a = ImageChops.lighter(a, r.getchannel("A"))
    a = a.point(lambda v: 0 if v <= 70 else 255 if v >= 190 else int((v - 70) * 255 / 120))
    resultado = original.convert("RGBA")
    resultado.putalpha(a.filter(ImageFilter.GaussianBlur(0.8)))
    return resultado


def recortar(img: Image.Image) -> Image.Image:
    # apaga restos quase transparentes do fundo
    a = img.getchannel("A").point(lambda v: 0 if v < 40 else v)
    img.putalpha(a)
    # caixa do sapato: ignora linhas/colunas com apenas alguns píxeis soltos
    solido = a.point(lambda v: 1 if v > 128 else 0)
    w, h = solido.size
    linhas = [sum(solido.crop((0, y, w, y + 1)).getdata()) for y in range(h)]
    colunas = [sum(solido.crop((x, 0, x + 1, h)).getdata()) for x in range(w)]
    minimo = 6
    ys = [i for i, v in enumerate(linhas) if v >= minimo]
    xs = [i for i, v in enumerate(colunas) if v >= minimo]
    if not ys or not xs:
        return img
    img = img.crop((xs[0], ys[0], xs[-1] + 1, ys[-1] + 1))
    w, h = img.size
    m = int(max(w, h) * MARGEM)
    tela = Image.new("RGBA", (w + 2 * m, h + 2 * m), (0, 0, 0, 0))
    tela.paste(img, (m, m))
    return tela


def main():
    todas = "--todas" in sys.argv
    # nomes de ficheiros específicos podem ser passados como argumentos (ex: 20_castanho.jpg)
    args = sys.argv[1:]
    refs = {args[i + 1] for i, a in enumerate(args[:-1]) if a == "--ref"}
    escolhidas = {a for i, a in enumerate(args) if not a.startswith("--") and (i == 0 or args[i - 1] != "--ref")}
    os.environ.setdefault("OMP_NUM_THREADS", "4")
    from rembg import new_session, remove

    ficheiros = sorted(f for f in ORIGEM.iterdir() if f.suffix.lower() in EXTENSOES)
    if escolhidas or refs:
        ficheiros = [f for f in ficheiros if f.name in escolhidas or (nome_destino(f) or ("",))[0] in refs]
        todas = True
    sessoes = {}
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
        modelo = MODELO_POR_REF.get(ref, MODELO)
        for m in modelo.split("+"):
            if m not in sessoes:
                sessoes[m] = new_session(m)
        img = ImageOps.exif_transpose(Image.open(f)).convert("RGB")
        img.thumbnail((ENTRADA_MAX, ENTRADA_MAX), Image.LANCZOS)
        if "+" in modelo:
            cortada = juntar_recortes(img, [remove(img, session=sessoes[m]) for m in modelo.split("+")])
        else:
            cortada = remove(img, session=sessoes[modelo])
            if modelo != MODELO:
                cortada = recuperar_sola(img, cortada)
        img = recortar(cortada)
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
