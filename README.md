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

### 3. Se for um modelo ou uma cor nova, acrescentar a linha no Excel

No `verao2027.xlsx`, uma linha por cor: Referência, Fornecedor, Numeração, Cor, Preço, Categoria.

### 4. Converter as fotos

```bash
python3 ferramentas/converter_fotos.py
```

- Só converte as fotos **novas ou alteradas** (as que já estavam feitas são ignoradas).
- Demora cerca de 5 segundos por foto. Da primeira vez descarrega o modelo de recorte (~180 MB), por isso demora mais.
- No fim aparece `Convertidas: X`. Cada foto dá origem a dois ficheiros na pasta `png/<referência>/`:
  a foto grande (`.png`) e uma miniatura leve (`mini/....webp`) usada na página inicial.
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
   **Todos os números** (a numeração do Excel, ex.: 35-41, 1 par de cada número, com quantas séries quiser)
   ou **Personalizado** (quantidade de cada número). O botão **Todas as cores · todos os números** escolhe tudo de uma vez.
2. Na página do **Carrinho** vê-se tudo, pode alterar-se ou remover cada modelo, e escrever o nome/loja e notas.
3. **Encomendar por WhatsApp** cria uma imagem da encomenda (foto de cada modelo com as cores e números)
   e abre o WhatsApp para o número de encomendas com a encomenda completa já escrita.
   O WhatsApp não deixa um site anexar ficheiros sozinho, por isso o cliente guarda/partilha a imagem e junta-a na conversa.

O carrinho fica guardado no browser de cada cliente (não precisa de conta).
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
