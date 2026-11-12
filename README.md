# Maneira Vaidosa — Catálogo de calçado

Site estático (HTML + CSS + JavaScript, sem instalação) que funciona em computador e telemóvel.
O site lê **diretamente** o Excel `verao2027.xlsx` e as fotos da pasta `png/` — não é preciso mexer no código para acrescentar modelos.

## Quando tiver fotos novas (passo a passo)

Abra o terminal na pasta do projeto (onde está o `index.html`) e siga estes passos.

### 0. Só da primeira vez: instalar o programa que tira o fundo às fotos

```bash
python3 -m pip install "rembg[cpu]" pillow
```

### 1. Trazer a versão mais recente do GitHub

```bash
git pull
```

### 2. Pôr as fotos na pasta `fotos originais`

O nome de cada foto tem de ser **referência + `_` + cor**, por exemplo:

| Foto | Aparece no site como |
|---|---|
| `20_castanho.jpg` | modelo 20, cor Castanho |
| `2709_camel.jpg` | modelo 2709, cor Camel |

- A cor tem de ser **igual à do Excel** (maiúsculas/minúsculas não importa; acentos sim).
- Para **substituir** uma foto, basta pôr a nova com o mesmo nome por cima da antiga.
- Não deixe nomes como `2720_taupe(1).jpg` (cópias): o `(1)` passaria a ser uma cor diferente.
- **Atenção:** pôr as fotos nesta pasta não chega — é preciso convertê-las (passo 4), senão não aparecem no site.

### 3. Se for um modelo ou uma cor nova, acrescentar a linha no Excel

No `verao2027.xlsx`, uma linha por cor: Referência, Fornecedor, Numeração, Numeração 2, Cor, Preço, Categoria, Sexo.

- **Numeração 2** pode ficar vazia quando o modelo só tem uma numeração.
- **Sexo**: Mulher, Homem ou Criança (aparecem como filtros no topo do site).
- Se a numeração for nova (ex.: `28-34`), acrescente-a ao Excel `Distribuiçao de Numeracoes.xlsx`.
- Se o mesmo sapato existir com outra referência (ex.: versão de criança), pode usar as mesmas fotos:
  acrescente-o a `CONFIG.mesmasFotos` no início de `app.js` (ex.: `"3521C": "3521"`, ou `"2703 Homem": "3521"`
  quando a referência também existe para outro sexo).

### 4. Converter as fotos

```bash
python3 ferramentas/converter_fotos.py
```

- Só converte as fotos **novas ou alteradas** (as que já estavam feitas são ignoradas).
- Demora cerca de 5 segundos por foto. Da primeira vez descarrega o modelo de recorte (~180 MB), por isso demora mais.
- No fim aparece `Convertidas: X`. Cada foto dá origem a dois ficheiros na pasta `png/<referência>/`:
  a foto grande (`.png`) e uma miniatura leve (`mini/....webp`) usada na página inicial.
- Se uma foto ficar com "nevoeiro" (a sombra do chão) por baixo da sola, acrescente a referência a
  `MODELO_POR_REF` no início de `ferramentas/converter_fotos.py` (usa o modelo `u2net`, melhor para solas de cortiça)
  e volte a converter essa referência: `python3 ferramentas/converter_fotos.py --ref 1128`
- Se uma foto ficar mal recortada, pode voltar a convertê-la sozinha:

  ```bash
  python3 ferramentas/converter_fotos.py 20_castanho.jpg
  ```

### 5. Ver o resultado no computador (opcional)

```bash
python3 -m http.server 8000
```

Abra http://localhost:8000 no browser. Para parar o servidor: `Ctrl + C`.

### 6. Enviar para o GitHub

```bash
git add -A
```

```bash
git commit -m "Fotos novas"
```

```bash
git push
```

Quando o site estiver publicado, as fotos e os modelos novos aparecem ao recarregar a página.

Notas:
- Uma categoria nova (ex.: "Bota") cria automaticamente um novo filtro no menu.
- Se o fornecedor estiver vazio, a referência aparece como `0-<referência>`.
- Se faltar a foto de uma cor, o cartão mostra outra cor do mesmo modelo; na página do modelo aparece "Foto brevemente".
- Cores novas recebem uma bolinha num tom automático; para escolher o tom exato, acrescente-a ao objeto `CORES` no início de `app.js`.
- Para ler mais ficheiros Excel (ex.: `inverno2027.xlsx`), acrescente-os a `CONFIG.ficheiros` em `app.js`.

## Encomendas por WhatsApp

As encomendas fazem-se pelo **carrinho** (ícone do saco, no topo):

1. Na página de cada modelo, **Adicionar ao carrinho** abre a escolha de cores. Para cada cor escolhe-se
   **Séries** (as numerações das colunas `Numeração` e `Numeração 2` do Excel, ex.: 35-40 e 36-41, com quantas de cada)
   ou **Personalizado** (quantidade de cada número, começa tudo a 0).
   Os botões **Todas as cores · série …** escolhem todas as cores de uma vez.
2. Na página do **Carrinho** vê-se tudo, pode alterar-se ou remover cada modelo, e escrever o nome/loja e notas.
3. **Enviar encomenda** cria uma imagem da encomenda (foto de cada modelo com as cores e números).
   No telemóvel, **Enviar** abre a partilha do telemóvel com a imagem já anexada: o cliente escolhe o WhatsApp
   e a conversa da Maneira Vaidosa. Onde a partilha não existe (alguns computadores), descarrega-se a imagem.

O carrinho fica guardado no browser de cada cliente (não precisa de conta).

### Quantos pares tem uma série

Está no Excel `Distribuiçao de Numeracoes.xlsx` (colunas **Numeração** e **Distribuição**), por exemplo:

| Numeração | Distribuição |
|---|---|
| 35-40 | 1 2 3 3 2 1 |
| 28-34 | 1 2 2 2 2 2 1 |

Ex.: 35-40 = 35×1, 36×2, 37×3, 38×3, 39×2, 40×1 (12 pares). Uma numeração que não esteja nesse Excel leva 1 par de cada número.
No carrinho e na imagem da encomenda aparece sempre a contagem de cada número.

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
