# Maneira Vaidosa — Catálogo de calçado

Site estático (HTML + CSS + JavaScript, sem instalação) que funciona em computador e telemóvel.
O site lê **diretamente** o Excel `verao2027.xlsx` e as fotos da pasta `png/` — não é preciso mexer no código para acrescentar modelos.

## Como acrescentar modelos

1. Acrescente linhas ao `verao2027.xlsx` (mesmas colunas: Referência, Fornecedor, Numeração, Cor, Preço, Categoria).
2. Ponha as fotos tiradas com a câmara na pasta `fotos originais/`, com o nome `<referência>_<cor>.jpg`
   (ex.: `20_castanho.jpg`).
3. Converta as fotos (só é preciso instalar uma vez):

   ```bash
   pip install "rembg[cpu]" pillow
   ```

   ```bash
   python ferramentas/converter_fotos.py
   ```

   O script tira o fundo, recorta o sapato e cria `png/<referência>/<referência> <cor>.png`
   (foto grande) e `png/<referência>/mini/<referência> <cor>.webp` (miniatura leve para a grelha).
   Só converte fotos novas ou alteradas; `--todas` volta a converter tudo.
4. Faça commit/push. Ao recarregar a página, os modelos novos aparecem.

Também pode pôr diretamente um PNG em `png/<referência>/<referência> <cor>.png` — o site usa-o mesmo sem miniatura.

Notas:
- Uma categoria nova (ex.: "Bota") cria automaticamente um novo filtro no menu.
- Se o fornecedor estiver vazio, a referência aparece como `0-<referência>`.
- Se faltar a foto de uma cor, o cartão mostra outra cor do mesmo modelo; na página do modelo aparece "Foto brevemente".
- Cores novas recebem uma bolinha num tom automático; para escolher o tom exato, acrescente-a ao objeto `CORES` no início de `app.js`.
- Para ler mais ficheiros Excel (ex.: `inverno2027.xlsx`), acrescente-os a `CONFIG.ficheiros` em `app.js`.

## Encomendas por WhatsApp

O botão **Encomendar** (página de cada modelo) abre o WhatsApp com uma mensagem para o número de encomendas,
já com a referência (`fornecedor-referência`), a categoria, a cor escolhida e o link da foto.
O número está em `CONFIG.empresa.whatsappEncomendas` no início de `app.js` e não aparece no site.

## Formatos

- Referência: `fornecedor-referência` → `1-20`
- Preço: `0` + preço em cêntimos com 4 dígitos + `0` → 7,50€ = `007500`, 10€ = `010000`, 16€ = `016000`

## Ver o site localmente

O site precisa de ser servido por um servidor (o browser não deixa ler o Excel se abrir o `index.html` com duplo clique):

```bash
python3 -m http.server 8000
```

e abrir http://localhost:8000

## Publicar

Basta ativar o **GitHub Pages** no repositório (Settings → Pages → Branch `main`, pasta `/root`).
