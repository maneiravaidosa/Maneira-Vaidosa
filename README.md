# Maneira Vaidosa — Catálogo de calçado

Site estático (HTML + CSS + JavaScript, sem instalação) que funciona em computador e telemóvel.
O site lê **diretamente** o Excel `verao2027.xlsx` e as fotos da pasta `png/` — não é preciso mexer no código para acrescentar modelos.

## Como acrescentar modelos

1. Acrescente linhas ao `verao2027.xlsx` (mesmas colunas: Referência, Fornecedor, Numeração, Cor, Preço, Categoria).
2. Crie a pasta `png/<referência>/` com uma foto por cor, chamada `<referência> <cor em minúsculas>.png`
   (ex.: `png/20/20 castanho.png`).
3. Faça commit/push. Ao recarregar a página, os modelos novos aparecem.

Notas:
- Uma categoria nova (ex.: "Bota") cria automaticamente um novo filtro no menu.
- Se o fornecedor estiver vazio, a referência aparece como `0-<referência>`.
- Se faltar a foto de uma cor, o cartão mostra outra cor do mesmo modelo; na página do modelo aparece "Foto brevemente".
- Cores novas recebem uma bolinha num tom automático; para escolher o tom exato, acrescente-a ao objeto `CORES` no início de `app.js`.
- Para ler mais ficheiros Excel (ex.: `inverno2027.xlsx`), acrescente-os a `CONFIG.ficheiros` em `app.js`.

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
