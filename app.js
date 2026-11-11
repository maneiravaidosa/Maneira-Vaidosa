/* =========================================================================
   Maneira Vaidosa — catálogo
   O site lê diretamente o(s) ficheiro(s) Excel e as fotos da pasta png/.
   Para acrescentar modelos basta:
     1. acrescentar linhas ao Excel (Referência, Fornecedor, Numeração, Cor, Preço, Categoria)
     2. criar a pasta png/<referência>/ com as fotos "<referência> <cor em minúsculas>.png"
   ========================================================================= */

const CONFIG = {
  // Ficheiros Excel a ler (podem acrescentar-se mais, ex: "inverno2027.xlsx")
  ficheiros: ["verao2027.xlsx"],
  colecao: "Verão 2027",
  pastaFotos: "png",
  // Quantos pares de cada número leva uma série, conforme o número de tamanhos da série.
  // Ex.: série 35-40 (6 números) → 35×1, 36×2, 37×3, 38×3, 39×2, 40×1 = 12 pares.
  // Séries com outro número de tamanhos que não esteja aqui levam 1 par de cada número.
  curvas: {
    6: [1, 2, 3, 3, 2, 1],
  },
  empresa: {
    nome: "Maneira Vaidosa",
    telefone: "256 812 118",
    morada: ["Travessa do Emigrante, Nº 120", "3700-475 Arrifana, Portugal"],
  },
};

// Cor das bolinhas. Cores novas que não estejam aqui recebem um tom neutro automático.
const CORES = {
  branco: "#ffffff",
  platina: "linear-gradient(135deg,#f5e7b8 0%,#d4b15f 45%,#f1dc9c 62%,#b8913f 100%)", // tom dourado
  taupe: "#9a8a78",
  castanho: "#6e4a2e",
  camel: "#c39a6b",
  marino: "#1f2b45",
  preto: "#1c1a18",
  verde: "#6f7a52",
  bege: "#e3d3b8",
  nude: "#e2bfa3",
  dourado: "linear-gradient(135deg,#f3e3a6,#c49a3c 55%,#f0d98a)",
  prata: "linear-gradient(135deg,#f4f4f4,#a9a9a9 55%,#e6e6e6)",
  azul: "#3c5a8a",
  vermelho: "#a3322b",
  rosa: "#e3a5a8",
  cinza: "#8d8d8d",
  amarelo: "#e5c35a",
  laranja: "#d77a3c",
};

/* ---------- utilitários ---------- */
const $ = (sel, el = document) => el.querySelector(sel);
const esc = (s) =>
  String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const semAcentos = (s) => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "");
const slug = (s) => semAcentos(s).toLowerCase().trim().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
const capitalizar = (s) => String(s).toLowerCase().replace(/(^|\s)\S/g, (c) => c.toUpperCase());
const plural = (s) => (/s$/i.test(s) ? s : s + "s");
const aleatorio = (arr) => arr[Math.floor(Math.random() * arr.length)];

function baralhar(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function lerPreco(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace("€", "").replace(",", ".").trim());
  return isNaN(n) ? 0 : n;
}

// 7.50€ -> 007500 · 10€ -> 010000 · 16€ -> 016000
function codigoPreco(preco) {
  const centimos = Math.round(preco * 100);
  return "0" + String(centimos).padStart(4, "0") + "0";
}

function corSwatch(cor) {
  const k = semAcentos(cor).toLowerCase().trim();
  if (CORES[k]) return CORES[k];
  // cor desconhecida: tom terroso derivado do nome
  let h = 0;
  for (const ch of k) h = (h * 31 + ch.charCodeAt(0)) % 360;
  return `hsl(${20 + (h % 30)}, ${25 + (h % 20)}%, ${40 + (h % 30)}%)`;
}

// formato "png" = foto grande · "mini" = miniatura leve (png/<ref>/mini/<ref> <cor>.webp) · "PNG" = extensão em maiúsculas
function caminhoFoto(ref, cor, formato = "png") {
  const nome = encodeURIComponent(`${ref} ${cor.toLowerCase().trim()}`);
  const pasta = `${CONFIG.pastaFotos}/${encodeURIComponent(ref)}`;
  return formato === "mini" ? `${pasta}/mini/${nome}.webp` : `${pasta}/${nome}.${formato}`;
}

/* ---------- dados ---------- */
const estado = {
  produtos: [], // [{ref, fornecedor, codigo, numeracao, preco, precoCodigo, categoria, catSlug, cores:[...], corInicial}]
  categorias: [], // [{nome, slug, total}]
};

function normalizarCabecalho(h) {
  return semAcentos(h).toLowerCase().replace(/[^a-z0-9]/g, "");
}

async function carregarExcel(ficheiro) {
  const resp = await fetch(ficheiro, { cache: "no-store" });
  if (!resp.ok) throw new Error(`Não foi possível abrir ${ficheiro}`);
  const wb = XLSX.read(await resp.arrayBuffer(), { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const linhas = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  return linhas.map((l) => {
    const o = {};
    for (const [k, v] of Object.entries(l)) o[normalizarCabecalho(k)] = v;
    return {
      ref: o.referencia != null ? String(o.referencia).trim() : "",
      fornecedor: o.fornecedor != null && String(o.fornecedor).trim() !== "" ? String(o.fornecedor).trim() : "0",
      numeracao: o.numeracao != null ? String(o.numeracao).trim() : "",
      numeracao2: o.numeracao2 != null ? String(o.numeracao2).trim() : "",
      cor: o.cor != null ? capitalizar(String(o.cor).trim()) : "",
      preco: lerPreco(o.preco),
      categoria: o.categoria != null ? capitalizar(String(o.categoria).trim()) : "Outros",
    };
  }).filter((l) => l.ref);
}

async function carregarDados() {
  const todas = (await Promise.all(CONFIG.ficheiros.map(carregarExcel))).flat();
  const mapa = new Map();
  for (const l of todas) {
    let p = mapa.get(l.ref);
    if (!p) {
      p = {
        ref: l.ref,
        fornecedor: l.fornecedor,
        codigo: `${l.fornecedor}-${l.ref}`,
        numeracao: l.numeracao,
        preco: l.preco,
        precoCodigo: codigoPreco(l.preco),
        categoria: l.categoria,
        catSlug: slug(l.categoria),
        cores: [],
        series: [], // ex.: ["35-40", "36-41"] (colunas Numeração e Numeração 2)
      };
      mapa.set(l.ref, p);
    }
    for (const sr of [l.numeracao, l.numeracao2]) if (sr && !p.series.includes(sr)) p.series.push(sr);
    if (l.cor && !p.cores.includes(l.cor)) p.cores.push(l.cor);
  }
  estado.produtos = [...mapa.values()].filter((p) => p.cores.length);
  for (const p of estado.produtos) {
    if (!p.series.length) p.series.push(p.numeracao || "Único");
    p.numeracao = p.series.join(" · ");
  }
  for (const p of estado.produtos) p.corInicial = aleatorio(p.cores); // muda a cada refresh

  const cats = new Map();
  for (const p of estado.produtos) {
    const c = cats.get(p.catSlug) || { nome: p.categoria, slug: p.catSlug, total: 0 };
    c.total++;
    cats.set(p.catSlug, c);
  }
  estado.categorias = [...cats.values()];
}

/* ---------- imagens com fallback ---------- */
// Se a foto de uma cor não existir, tenta .PNG; depois tenta as outras cores; por fim mostra um marcador.
// Imagem com alternativas: se a miniatura não existir usa a foto grande; se a cor não tiver foto, tenta outra cor.
function htmlImagem(p, cor, extraClass = "", mini = true) {
  const ordem = [cor, ...p.cores.filter((c) => c !== cor)];
  const formatos = mini ? "mini|png|PNG" : "png|PNG";
  return `<img class="shoe ${extraClass}" src="${caminhoFoto(p.ref, cor, formatos.split("|")[0])}" alt="Modelo ${esc(p.codigo)} — ${esc(cor)}"
    loading="lazy" decoding="async" data-ref="${esc(p.ref)}" data-ordem="${esc(ordem.join("|"))}" data-i="0"
    data-formatos="${formatos}" data-f="0">`;
}

function trocarFoto(img, ref, cor) {
  img.dataset.ordem = cor;
  img.dataset.i = "0";
  img.dataset.f = "0";
  img.closest(".shoe-frame")?.classList.remove("is-missing");
  img.src = caminhoFoto(ref, cor, img.dataset.formatos.split("|")[0]);
}

document.addEventListener(
  "error",
  (e) => {
    const img = e.target;
    if (!(img instanceof HTMLImageElement) || !img.dataset.ordem) return;
    const ordem = img.dataset.ordem.split("|");
    const formatos = img.dataset.formatos.split("|");
    let i = +img.dataset.i;
    let f = +img.dataset.f + 1;
    if (f >= formatos.length) {
      // nenhum formato desta cor existe: passa para a cor seguinte (exceto na página do modelo)
      if (img.dataset.fixa === "1" || i + 1 >= ordem.length) {
        img.removeAttribute("data-ordem");
        img.closest(".shoe-frame")?.classList.add("is-missing");
        img.src = "data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==";
        return;
      }
      img.dataset.i = ++i;
      f = 0;
    }
    img.dataset.f = f;
    img.src = caminhoFoto(img.dataset.ref, ordem[i], formatos[f]);
  },
  true
);

/* ---------- componentes ---------- */
function htmlDots(p, ativa, grande = false) {
  return p.cores
    .map((c) => {
      const attrs = `class="dot ${grande ? "dot--lg" : ""} ${c === ativa ? "is-active" : ""}" style="--sw:${corSwatch(c)}" data-cor="${esc(c)}" title="${esc(c)}"`;
      // nos cartões (dentro de um link) usa-se <span>; na página do produto, botões acessíveis
      return grande
        ? `<button type="button" ${attrs} aria-label="Cor ${esc(c)}" aria-pressed="${c === ativa}"></button>`
        : `<span ${attrs}></span>`;
    })
    .join("");
}

function htmlCard(p, i) {
  return `
  <a class="card reveal" href="#/p/${encodeURIComponent(p.ref)}" style="--d:${(i % 4) * 70}ms" data-ref="${esc(p.ref)}">
    <div class="card__media shoe-frame">
      ${htmlImagem(p, p.corInicial)}
      <span class="card__tag">${esc(p.categoria)}</span>
    </div>
    <div class="card__body">
      <div class="card__row">
        <span class="card__ref">${esc(p.codigo)}</span>
        <span class="card__price">${esc(p.precoCodigo)}</span>
      </div>
      <div class="card__row card__row--sub">
        <span class="card__dots">${htmlDots(p, p.corInicial)}</span>
        <span class="card__sizes">${esc(p.numeracao)}</span>
      </div>
    </div>
  </a>`;
}

function htmlNav(ativa) {
  const links = [
    `<a href="#/" class="nav__link ${ativa === "todos" ? "is-active" : ""}">Todos</a>`,
    ...estado.categorias.map(
      (c) => `<a href="#/c/${c.slug}" class="nav__link ${ativa === c.slug ? "is-active" : ""}">${esc(plural(c.nome))}</a>`
    ),
    `<a href="#/contactos" class="nav__link nav__link--contact ${ativa === "contactos" ? "is-active" : ""}">Contactos</a>`,
  ];
  return links.join("");
}

function htmlGrelha(lista) {
  if (!lista.length) return `<p class="empty">Ainda não há modelos nesta categoria.</p>`;
  return `<div class="grid">${lista.map(htmlCard).join("")}</div>`;
}

function htmlChips(ativa) {
  return `<div class="chips">
    <a href="#/" class="chip ${ativa === "todos" ? "is-active" : ""}">Todos <sup>${estado.produtos.length}</sup></a>
    ${estado.categorias
      .map(
        (c) =>
          `<a href="#/c/${c.slug}" class="chip ${ativa === c.slug ? "is-active" : ""}">${esc(plural(c.nome))} <sup>${c.total}</sup></a>`
      )
      .join("")}
  </div>`;
}

/* ---------- páginas ---------- */
function paginaInicio() {
  const destaque = baralhar(estado.produtos).slice(0, 5);
  const cats = estado.categorias.map((c) => plural(c.nome));
  const marquee = Array(4).fill(cats.map((c) => `<span>${esc(c)}</span><i>✦</i>`).join("")).join("");

  return `
  <section class="hero">
    <div class="hero__bg" aria-hidden="true"><div class="blob blob--1"></div><div class="blob blob--2"></div></div>
    <div class="hero__floats" aria-hidden="true">
      ${destaque
        .map(
          (p, i) => `<a class="float float--${i + 1}" href="#/p/${encodeURIComponent(p.ref)}" tabindex="-1" data-depth="${(i % 3) + 1}">
            <span class="shoe-frame">${htmlImagem(p, p.corInicial)}</span></a>`
        )
        .join("")}
    </div>
    <div class="hero__content">
      <p class="eyebrow"><span class="eyebrow__line"></span>Coleção ${esc(CONFIG.colecao)}</p>
      <h1 class="hero__title">
        <span class="line"><span>Maneira</span></span>
        <span class="line"><em>Vaidosa</em></span>
      </h1>
      <p class="hero__lead">Calçado para Verão 2027 já disponível!</p>
      <div class="hero__cta">
        <a href="#colecao" class="btn" data-scroll>Ver coleção</a>
        <a href="#/contactos" class="btn btn--ghost">Contactos</a>
      </div>
    </div>
    <a href="#colecao" class="scroll-hint" data-scroll aria-label="Descer até à coleção"><span>Scroll</span><i></i></a>
  </section>

  <div class="marquee" aria-hidden="true"><div class="marquee__track">${marquee}</div></div>

  <section class="catalog" id="colecao">
    <div class="catalog__head">
      <div>
        <p class="eyebrow"><span class="eyebrow__line"></span>${estado.produtos.length} modelos</p>
        <h2 class="section-title">A <em>coleção</em></h2>
      </div>
      ${htmlChips("todos")}
    </div>
    ${htmlGrelha(estado.produtos)}
  </section>`;
}

function paginaCategoria(catSlug) {
  const cat = estado.categorias.find((c) => c.slug === catSlug);
  if (!cat) return pagina404();
  const lista = estado.produtos.filter((p) => p.catSlug === catSlug);
  return `
  <section class="page-head">
    <p class="eyebrow"><span class="eyebrow__line"></span>${esc(CONFIG.colecao)} · ${lista.length} modelos</p>
    <h1 class="page-title">${esc(plural(cat.nome))}</h1>
  </section>
  <section class="catalog catalog--flush">
    <div class="catalog__head catalog__head--right">${htmlChips(catSlug)}</div>
    ${htmlGrelha(lista)}
  </section>`;
}

function paginaProduto(ref) {
  const p = estado.produtos.find((x) => x.ref === ref);
  if (!p) return pagina404();
  const idx = estado.produtos.indexOf(p);
  const ant = estado.produtos[(idx - 1 + estado.produtos.length) % estado.produtos.length];
  const seg = estado.produtos[(idx + 1) % estado.produtos.length];
  const semelhantes = baralhar(estado.produtos.filter((x) => x.catSlug === p.catSlug && x !== p)).slice(0, 4);

  const tamanhos = tamanhosDe(p).map((n) => `<span class="size">${esc(n)}</span>`).join("");

  return `
  <section class="product" data-ref="${esc(p.ref)}">
    <div class="product__crumbs">
      <a href="#/">Coleção</a><span>/</span><a href="#/c/${p.catSlug}">${esc(plural(p.categoria))}</a><span>/</span><b>${esc(p.codigo)}</b>
    </div>
    <div class="product__grid">
      <div class="product__media">
        <div class="stage shoe-frame" id="stage">
          <div class="stage__ring" aria-hidden="true"></div>
          ${htmlImagem(p, p.corInicial, "stage__img", false)}
          <span class="stage__missing">Foto brevemente</span>
        </div>
      </div>
      <div class="product__info">
        <p class="eyebrow"><span class="eyebrow__line"></span>${esc(p.categoria)} · ${esc(CONFIG.colecao)}</p>
        <h1 class="product__ref">${esc(p.codigo)}</h1>
        <dl class="specs">
          <div><dt>Referência</dt><dd>${esc(p.codigo)}</dd></div>
          <div><dt>Preço</dt><dd class="mono">${esc(p.precoCodigo)}</dd></div>
          <div><dt>Categoria</dt><dd>${esc(p.categoria)}</dd></div>
          <div><dt>Numeração</dt><dd>${esc(p.numeracao)}</dd></div>
        </dl>
        <div class="block block--cores">
          <div class="block__label">Cor <span id="corNome">${esc(p.corInicial)}</span></div>
          <div class="dots" id="dots">${htmlDots(p, p.corInicial, true)}</div>
        </div>
        <div class="block">
          <div class="block__label">Tamanhos</div>
          <div class="sizes">${tamanhos}</div>
        </div>
        <div class="product__actions">
          <button type="button" class="btn" id="addCarrinho">
            <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round" d="M5 8h14l-1.2 12.2a1 1 0 0 1-1 .8H7.2a1 1 0 0 1-1-.8L5 8Z"/><path fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" d="M9 10V7a3 3 0 0 1 6 0v3"/></svg>
            <span id="addCarrinhoTxt">Adicionar ao carrinho</span>
          </button>
          <a class="btn btn--ghost" href="#/contactos">Contactos</a>
        </div>
        <p class="product__nocart" id="noCarrinho"></p>
        <div class="product__pager">
          <a href="#/p/${encodeURIComponent(ant.ref)}">← ${esc(ant.codigo)}</a>
          <a href="#/p/${encodeURIComponent(seg.ref)}">${esc(seg.codigo)} →</a>
        </div>
      </div>
    </div>
  </section>
  ${
    semelhantes.length
      ? `<section class="catalog catalog--related">
          <div class="catalog__head"><h2 class="section-title">Também em <em>${esc(plural(p.categoria).toLowerCase())}</em></h2></div>
          ${htmlGrelha(semelhantes)}
        </section>`
      : ""
  }`;
}

function paginaContactos() {
  const e = CONFIG.empresa;
  const mapa = "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(e.morada.join(", "));
  return `
  <section class="contact">
    <div class="contact__bg" aria-hidden="true"><div class="blob blob--1"></div></div>
    <p class="eyebrow"><span class="eyebrow__line"></span>Contactos</p>
    <h1 class="contact__title"><span class="line"><span>Maneira</span></span><span class="line"><em>Vaidosa</em></span></h1>
    <div class="contact__cards">
      <a class="contact__card reveal" href="tel:+351256812118">
        <span class="contact__label">Telefone</span>
        <span class="contact__value">${esc(e.telefone)}</span>
        <span class="contact__more">Ligar →</span>
      </a>
      <a class="contact__card reveal" style="--d:90ms" href="${mapa}" target="_blank" rel="noopener">
        <span class="contact__label">Morada</span>
        <span class="contact__value contact__value--sm">${e.morada.map(esc).join("<br>")}</span>
        <span class="contact__more">Ver no mapa →</span>
      </a>
    </div>
  </section>`;
}

function pagina404() {
  return `<section class="page-head page-head--center">
    <p class="eyebrow">Ups</p>
    <h1 class="page-title">Página não encontrada</h1>
    <p><a class="btn" href="#/">Voltar à coleção</a></p>
  </section>`;
}

/* ---------- interações ---------- */
let revealObserver;
function ativarReveal() {
  revealObserver?.disconnect();
  const els = document.querySelectorAll(".reveal");
  if (!("IntersectionObserver" in window)) return els.forEach((el) => el.classList.add("is-in"));
  revealObserver = new IntersectionObserver(
    (entries) => {
      for (const en of entries) {
        if (en.isIntersecting) {
          en.target.classList.add("is-in");
          revealObserver.unobserve(en.target);
        }
      }
    },
    { rootMargin: "0px 0px -8% 0px", threshold: 0.08 }
  );
  els.forEach((el) => revealObserver.observe(el));
}

function ativarProduto(ref) {
  const p = estado.produtos.find((x) => x.ref === ref);
  const dots = $("#dots");
  if (!p || !dots) return;
  const stage = $("#stage");
  const img = $(".stage__img", stage);
  dots.addEventListener("click", (e) => {
    const b = e.target.closest(".dot");
    if (!b || b.classList.contains("is-active")) return;
    const cor = b.dataset.cor;
    dots.querySelectorAll(".dot").forEach((d) => {
      d.classList.toggle("is-active", d === b);
      d.setAttribute("aria-pressed", d === b);
    });
    $("#corNome").textContent = cor;
    stage.classList.remove("is-missing");
    img.classList.add("is-swapping");
    setTimeout(() => {
      img.dataset.fixa = "1"; // na página do produto, a cor escolhida não troca por outra
      img.alt = `Modelo ${p.codigo} — ${cor}`;
      trocarFoto(img, p.ref, cor);
      const done = () => img.classList.remove("is-swapping");
      if (img.complete) requestAnimationFrame(done);
      else img.addEventListener("load", done, { once: true }), img.addEventListener("error", done, { once: true });
    }, 180);
  });

  $("#addCarrinho").addEventListener("click", () => abrirEscolha(p, $("#dots .is-active")?.dataset.cor));
  atualizarBotaoProduto(p);

  // efeito 3D suave ao passar o rato
  if (window.matchMedia("(hover: hover)").matches) {
    stage.addEventListener("pointermove", (e) => {
      const r = stage.getBoundingClientRect();
      const x = (e.clientX - r.left) / r.width - 0.5;
      const y = (e.clientY - r.top) / r.height - 0.5;
      stage.style.setProperty("--rx", `${(-y * 8).toFixed(2)}deg`);
      stage.style.setProperty("--ry", `${(x * 10).toFixed(2)}deg`);
    });
    stage.addEventListener("pointerleave", () => {
      stage.style.setProperty("--rx", "0deg");
      stage.style.setProperty("--ry", "0deg");
    });
  }
}

// Nos cartões: clicar (ou tocar) numa bolinha muda a cor sem abrir a página do modelo;
// no computador, passar o rato por cima também pré-visualiza. Ao abrir o modelo, abre nessa cor.
function ativarCards() {
  document.querySelectorAll(".card").forEach((card) => {
    const img = $(".shoe", card);
    const p = estado.produtos.find((x) => x.ref === card.dataset.ref);
    const escolher = (d) => {
      if (d.classList.contains("is-active")) return;
      card.querySelectorAll(".dot").forEach((x) => x.classList.toggle("is-active", x === d));
      trocarFoto(img, card.dataset.ref, d.dataset.cor);
      if (p) p.corInicial = d.dataset.cor;
    };
    card.addEventListener("pointerover", (e) => {
      const d = e.target.closest(".dot");
      if (d && e.pointerType === "mouse") escolher(d);
    });
    card.addEventListener("click", (e) => {
      const d = e.target.closest(".dot");
      if (!d) return;
      e.preventDefault();
      e.stopPropagation();
      escolher(d);
    });
  });
}

let parallaxRAF;
function ativarHero() {
  const hero = $(".hero");
  if (!hero || window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
  const floats = [...hero.querySelectorAll(".float")];
  let mx = 0, my = 0;
  hero.addEventListener("pointermove", (e) => {
    mx = e.clientX / window.innerWidth - 0.5;
    my = e.clientY / window.innerHeight - 0.5;
  });
  cancelAnimationFrame(parallaxRAF);
  const tick = () => {
    if (!document.body.contains(hero)) return;
    const sy = Math.min(window.scrollY, window.innerHeight);
    for (const f of floats) {
      const d = +f.dataset.depth;
      f.style.setProperty("--px", `${(mx * 18 * d).toFixed(1)}px`);
      f.style.setProperty("--py", `${(my * 14 * d - sy * 0.12 * d).toFixed(1)}px`);
    }
    parallaxRAF = requestAnimationFrame(tick);
  };
  tick();
}

/* ---------- carrinho ---------- */
// Fica guardado no próprio browser de quem encomenda (não precisa de conta). Formato:
// { itens: { "<ref>": { "<Cor>": { series: { "35-40": 1, "36-41": 2 } } | { pares: { "35": 2, "37": 1 } } } }, nome, notas }
//   series = quantas séries de cada numeração (ver CONFIG.curvas)
//   pares  = números personalizados, com a quantidade de cada número
const CHAVE_CARRINHO = "mv-carrinho";
const produtoPorRef = (ref) => estado.produtos.find((p) => p.ref === ref);

function intervalo(serie) {
  const m = String(serie).match(/^(\d+)\s*[-–a]\s*(\d+)$/);
  if (!m) return [String(serie)];
  const out = [];
  for (let n = +m[1]; n <= +m[2]; n++) out.push(String(n));
  return out;
}

// [["35",1],["36",2],["37",3],["38",3],["39",2],["40",1]]
function curvaSerie(serie) {
  const t = intervalo(serie);
  const c = CONFIG.curvas[t.length] || [];
  return t.map((n, i) => [n, c[i] ?? 1]);
}
const paresSerie = (serie) => curvaSerie(serie).reduce((a, [, q]) => a + q, 0);

// todos os números das séries do modelo (ex.: 35-40 e 36-41 → 35 a 41)
function tamanhosDe(p) {
  const todos = [...new Set(p.series.flatMap(intervalo))];
  return todos.sort((a, b) => (isNaN(a) || isNaN(b) ? 0 : a - b));
}

function paresLinha(p, l) {
  if (!l) return 0;
  if (l.series) return Object.entries(l.series).reduce((a, [sr, n]) => a + (+n || 0) * paresSerie(sr), 0);
  return Object.values(l.pares || {}).reduce((a, b) => a + (+b || 0), 0);
}

// "Série 35-40" · "Série 36-41 × 2" · "35×2 · 36×3"
function descreverLinha(p, l, sep = " · ") {
  if (l.series)
    return p.series
      .filter((sr) => l.series[sr] > 0)
      .map((sr) => `Série ${sr}${l.series[sr] > 1 ? ` × ${l.series[sr]}` : ""}`)
      .join(sep);
  return tamanhosDe(p)
    .filter((t) => l.pares?.[t] > 0)
    .map((t) => `${t}×${l.pares[t]}`)
    .join(sep);
}

const carrinho = {
  itens: {},
  nome: "",
  notas: "",
  carregar() {
    try {
      const d = JSON.parse(localStorage.getItem(CHAVE_CARRINHO) || "{}");
      this.nome = d.nome || "";
      this.notas = d.notas || "";
      this.itens = {};
      // ignora modelos/cores que entretanto saíram do Excel
      for (const [ref, cores] of Object.entries(d.itens || {})) {
        const p = produtoPorRef(ref);
        if (!p) continue;
        const ok = {};
        for (let [cor, l] of Object.entries(cores || {})) {
          if (l?.serie) l = { series: { [p.series[0]]: l.serie } }; // formato antigo
          if (l?.series) l = { series: Object.fromEntries(Object.entries(l.series).filter(([sr]) => p.series.includes(sr))) };
          if (p.cores.includes(cor) && paresLinha(p, l) > 0) ok[cor] = l;
        }
        if (Object.keys(ok).length) this.itens[ref] = ok;
      }
    } catch {
      this.itens = {};
    }
  },
  guardar() {
    try {
      localStorage.setItem(CHAVE_CARRINHO, JSON.stringify({ itens: this.itens, nome: this.nome, notas: this.notas }));
    } catch {}
  },
  // substitui as cores escolhidas para um modelo (vazio = tira o modelo do carrinho)
  definir(ref, cores) {
    if (Object.keys(cores).length) this.itens[ref] = cores;
    else delete this.itens[ref];
    this.guardar();
    atualizarContador(true);
  },
  esvaziar() {
    this.itens = {};
    this.guardar();
    atualizarContador();
  },
  totais() {
    let modelos = 0, cores = 0, pares = 0;
    for (const [ref, linhas] of Object.entries(this.itens)) {
      const p = produtoPorRef(ref);
      modelos++;
      for (const l of Object.values(linhas)) (cores++, (pares += paresLinha(p, l)));
    }
    return { modelos, cores, pares };
  },
};

const plur = (n, um, varios) => `${n} ${n === 1 ? um : varios}`;

function atualizarContador(bump = false) {
  const n = Object.keys(carrinho.itens).length;
  const el = $("#cartCount");
  if (!el) return;
  el.textContent = n;
  el.hidden = n === 0;
  $("#cartBtn").setAttribute("aria-label", `Carrinho (${plur(n, "modelo", "modelos")})`);
  if (bump) {
    const b = $("#cartBtn");
    b.classList.remove("is-bump");
    void b.offsetWidth;
    b.classList.add("is-bump");
  }
}

/* ---------- painéis (janela no computador, folha que sobe no telemóvel) ---------- */
function abrirPainel(html, classe = "") {
  fecharPainel();
  $(".toast")?.remove();
  const el = document.createElement("div");
  el.className = `painel ${classe}`;
  el.innerHTML = `<div class="painel__fundo" data-fechar></div><div class="painel__caixa" role="dialog" aria-modal="true">${html}</div>`;
  document.body.appendChild(el);
  document.documentElement.classList.add("sem-scroll");
  el.addEventListener("click", (e) => e.target.closest("[data-fechar]") && fecharPainel());
  requestAnimationFrame(() => el.classList.add("is-open"));
  return el;
}

function fecharPainel() {
  document.querySelectorAll(".painel").forEach((el) => el.remove());
  document.documentElement.classList.remove("sem-scroll");
}
document.addEventListener("keydown", (e) => e.key === "Escape" && fecharPainel());

function aviso(html) {
  $(".toast")?.remove();
  const t = document.createElement("div");
  t.className = "toast";
  t.setAttribute("role", "status");
  t.innerHTML = html;
  document.body.appendChild(t);
  requestAnimationFrame(() => t.classList.add("is-in"));
  setTimeout(() => {
    t.classList.remove("is-in");
    setTimeout(() => t.remove(), 400);
  }, 3200);
}

/* ---------- escolher cores e números ---------- */
const htmlPassos = (k, n, min = 0) => `<span class="step">
  <button type="button" data-k="${k}" data-d="-1" aria-label="Menos" ${n <= min ? "disabled" : ""}>−</button>
  <b>${n}</b>
  <button type="button" data-k="${k}" data-d="1" aria-label="Mais">+</button>
</span>`;

function abrirEscolha(p, corDestaque) {
  const tams = tamanhosDe(p);
  const jaNoCarrinho = !!carrinho.itens[p.ref];
  const sel = JSON.parse(JSON.stringify(carrinho.itens[p.ref] || {}));
  const novaSerie = (sr = p.series[0]) => ({ series: { [sr]: 1 } });
  // primeira vez: a cor que está a ser vista já vem escolhida com a primeira série
  if (!jaNoCarrinho && corDestaque) sel[corDestaque] = novaSerie();

  const htmlCor = (cor) => {
    const l = sel[cor];
    const modo = l ? (l.series ? "serie" : "pares") : "";
    const n = paresLinha(p, l);
    return `<div class="esc__cor ${n ? "is-on" : ""}" data-cor="${esc(cor)}">
      <div class="esc__top">
        <span class="dot dot--md" style="--sw:${corSwatch(cor)}"></span>
        <span class="esc__nome">${esc(cor)}</span>
        <span class="esc__pares">${n ? plur(n, "par", "pares") : ""}</span>
      </div>
      <div class="seg" role="group" aria-label="Números para a cor ${esc(cor)}">
        <button type="button" data-modo="" class="${modo === "" ? "is-on" : ""}">Nenhum</button>
        <button type="button" data-modo="serie" class="${modo === "serie" ? "is-on" : ""}">${p.series.length > 1 ? "Séries" : "Série"}</button>
        <button type="button" data-modo="pares" class="${modo === "pares" ? "is-on" : ""}">Personalizado</button>
      </div>
      ${
        modo === "serie"
          ? p.series
              .map((sr) => {
                const curva = curvaSerie(sr);
                const q = l.series[sr] || 0;
                return `<div class="esc__serie ${q ? "is-on" : ""}">
                  <span>Série ${esc(sr)} <small>${paresSerie(sr)} pares: ${curva.map(([t, c]) => `${t}×${c}`).join(" ")}</small></span>
                  ${htmlPassos("s:" + sr, q)}
                </div>`;
              })
              .join("")
          : ""
      }
      ${
        modo === "pares"
          ? `<div class="esc__tams">${tams
              .map((t) => {
                const q = l.pares[t] || 0;
                return `<div class="tam ${q ? "is-on" : ""}"><span class="tam__n">${esc(t)}</span>${htmlPassos("t:" + t, q)}</div>`;
              })
              .join("")}</div>`
          : ""
      }
    </div>`;
  };

  const painel = abrirPainel(
    `<header class="painel__topo">
      <div><p class="eyebrow">${esc(p.categoria)} · ${plur(p.cores.length, "cor", "cores")}</p><h2 class="painel__titulo">${esc(p.codigo)}</h2></div>
      <button type="button" class="painel__x" data-fechar aria-label="Fechar">×</button>
    </header>
    <div class="painel__corpo">
      <div class="esc__rapido">
        ${p.series
          .map((sr) => `<button type="button" class="chip chip--btn" data-todas="${esc(sr)}">Todas as cores · série ${esc(sr)}</button>`)
          .join("")}
        <button type="button" class="chip chip--btn" data-limpar>Limpar</button>
      </div>
      <div class="esc__lista">${p.cores.map(htmlCor).join("")}</div>
    </div>
    <footer class="painel__rodape">
      <div class="esc__total" id="escTotal"></div>
      <button type="button" class="btn" id="escOk"></button>
    </footer>`,
    "painel--escolha"
  );

  const lista = $(".esc__lista", painel);
  const redesenhar = (cor) => {
    const row = lista.querySelector(`.esc__cor[data-cor="${CSS.escape(cor)}"]`);
    row.outerHTML = htmlCor(cor);
    resumo();
  };
  const resumo = () => {
    let pares = 0, cores = 0;
    for (const cor of p.cores) {
      const n = paresLinha(p, sel[cor]);
      if (n) (cores++, (pares += n));
    }
    $("#escTotal").innerHTML = pares
      ? `<b>${pares}</b> ${pares === 1 ? "par" : "pares"} · ${plur(cores, "cor", "cores")}`
      : `Escolhe pelo menos uma cor`;
    const ok = $("#escOk");
    ok.disabled = !pares && !jaNoCarrinho;
    const resto = '<span class="so-largo"> carrinho</span>';
    ok.innerHTML = !pares && jaNoCarrinho ? "Tirar do carrinho" : jaNoCarrinho ? `Atualizar${resto}` : `Adicionar<span class="so-largo"> ao carrinho</span>`;
  };
  resumo();

  painel.addEventListener("click", (e) => {
    const row = e.target.closest(".esc__cor");
    const cor = row?.dataset.cor;
    const modoBtn = e.target.closest("[data-modo]");
    const passo = e.target.closest("[data-k]");
    if (modoBtn && cor) {
      const modo = modoBtn.dataset.modo;
      const antes = sel[cor];
      if (modo === "") delete sel[cor];
      else if (modo === "serie") sel[cor] = antes?.series ? antes : novaSerie();
      else sel[cor] = antes?.pares ? antes : { pares: Object.fromEntries(tams.map((t) => [t, 0])) }; // começa tudo a 0
      redesenhar(cor);
    } else if (passo && cor) {
      const d = +passo.dataset.d;
      const k = passo.dataset.k;
      if (k.startsWith("s:")) {
        const sr = k.slice(2);
        sel[cor].series[sr] = Math.max(0, (sel[cor].series[sr] || 0) + d);
      } else {
        const t = k.slice(2);
        sel[cor].pares[t] = Math.max(0, (sel[cor].pares[t] || 0) + d);
      }
      redesenhar(cor);
    } else if (e.target.closest("[data-todas]")) {
      const sr = e.target.closest("[data-todas]").dataset.todas;
      for (const c of p.cores) if (!paresLinha(p, sel[c])) sel[c] = novaSerie(sr);
      p.cores.forEach(redesenhar);
    } else if (e.target.closest("[data-limpar]")) {
      for (const c of p.cores) delete sel[c];
      p.cores.forEach(redesenhar);
    } else if (e.target.closest("#escOk")) {
      // guarda pela ordem das cores do Excel e só as que têm pares
      const final = {};
      for (const c of p.cores) {
        const l = sel[c];
        if (!paresLinha(p, l)) continue;
        const soPositivos = (o) => Object.fromEntries(Object.entries(o).filter(([, q]) => q > 0));
        final[c] = l.series ? { series: soPositivos(l.series) } : { pares: soPositivos(l.pares) };
      }
      carrinho.definir(p.ref, final);
      fecharPainel();
      aviso(
        Object.keys(final).length
          ? `<span>${esc(p.codigo)} no carrinho</span><a href="#/carrinho">Ver carrinho</a>`
          : `<span>${esc(p.codigo)} saiu do carrinho</span>`
      );
      atualizarBotaoProduto(p);
      redesenharCarrinho();
    }
  });
}

// Na página do modelo: diz o que já está no carrinho
function atualizarBotaoProduto(p) {
  const txt = $("#addCarrinhoTxt");
  if (!txt || $("#stage")?.closest(".product")?.dataset.ref !== p.ref) return;
  const linhas = carrinho.itens[p.ref];
  txt.textContent = linhas ? "Alterar no carrinho" : "Adicionar ao carrinho";
  const n = linhas ? Object.values(linhas).reduce((a, l) => a + paresLinha(p, l), 0) : 0;
  $("#noCarrinho").innerHTML = linhas
    ? `<span class="nocart__ok" aria-hidden="true">✓</span> No carrinho: ${plur(n, "par", "pares")} em ${plur(Object.keys(linhas).length, "cor", "cores")} · <a href="#/carrinho">Ver carrinho</a>`
    : "";
}

/* ---------- página do carrinho ---------- */
function htmlItemCarrinho(ref) {
  const p = produtoPorRef(ref);
  const linhas = Object.entries(carrinho.itens[ref]);
  const n = linhas.reduce((a, [, l]) => a + paresLinha(p, l), 0);
  return `<article class="citem" data-ref="${esc(ref)}">
    <a class="citem__foto shoe-frame" href="#/p/${encodeURIComponent(ref)}">${htmlImagem(p, linhas[0][0])}</a>
    <div class="citem__info">
      <div class="citem__topo">
        <a class="citem__ref" href="#/p/${encodeURIComponent(ref)}">${esc(p.codigo)}</a>
        <span class="card__price">${esc(p.precoCodigo)}</span>
      </div>
      <p class="citem__cat">${esc(p.categoria)} · ${plur(n, "par", "pares")}</p>
      <ul class="citem__linhas">
        ${linhas
          .map(
            ([cor, l]) => `<li>
              <span class="dot dot--static" style="--sw:${corSwatch(cor)}"></span>
              <b>${esc(cor)}</b>
              <span class="citem__pares">${plur(paresLinha(p, l), "par", "pares")}</span>
              <span class="citem__desc">${esc(descreverLinha(p, l))}</span>
            </li>`
          )
          .join("")}
      </ul>
      <div class="citem__acoes">
        <button type="button" class="link-btn" data-editar>Alterar</button>
        <button type="button" class="link-btn link-btn--muted" data-remover>Remover</button>
      </div>
    </div>
  </article>`;
}

const ICONE_ENVIAR = `<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round" d="M21 3 10.5 13.5M21 3l-6.5 18-4-7.5L3 9.5 21 3Z"/></svg>`;

function paginaCarrinho() {
  const refs = Object.keys(carrinho.itens);
  if (!refs.length)
    return `<section class="page-head page-head--center">
      <p class="eyebrow">Carrinho</p>
      <h1 class="page-title">O carrinho está vazio</h1>
      <p class="muted">Abre um modelo e carrega em <b>Adicionar ao carrinho</b>.</p>
      <p><a class="btn" href="#/">Ver coleção</a></p>
    </section>`;
  const t = carrinho.totais();
  return `
  <section class="page-head page-head--cart">
    <p class="eyebrow"><span class="eyebrow__line"></span>${esc(CONFIG.colecao)}</p>
    <h1 class="page-title">Carrinho</h1>
  </section>
  <section class="cart">
    <div class="cart__lista">${refs.map(htmlItemCarrinho).join("")}</div>
    <aside class="cart__resumo">
      <h2>Encomenda</h2>
      <dl class="cart__totais">
        <div><dt>Modelos</dt><dd>${t.modelos}</dd></div>
        <div><dt>Cores</dt><dd>${t.cores}</dd></div>
        <div><dt>Pares</dt><dd>${t.pares}</dd></div>
      </dl>
      <label class="campo"><span>Nome / loja</span><input id="cartNome" autocomplete="organization" placeholder="Para sabermos de quem é a encomenda" value="${esc(carrinho.nome)}"></label>
      <label class="campo"><span>Notas <i>(opcional)</i></span><textarea id="cartNotas" rows="2" placeholder="Ex.: prazo de entrega, morada…">${esc(carrinho.notas)}</textarea></label>
      <button type="button" class="btn btn--full" id="encomendar">${ICONE_ENVIAR}<span>Enviar encomenda</span></button>
      <button type="button" class="link-btn link-btn--muted" id="esvaziar">Esvaziar carrinho</button>
    </aside>
  </section>`;
}

function redesenharCarrinho() {
  if (document.body.dataset.page !== "carrinho") return;
  $("#app").innerHTML = paginaCarrinho();
  ativarCarrinho();
}

function ativarCarrinho() {
  const app = $("#app");
  $("#cartNome")?.addEventListener("input", (e) => ((carrinho.nome = e.target.value), carrinho.guardar()));
  $("#cartNotas")?.addEventListener("input", (e) => ((carrinho.notas = e.target.value), carrinho.guardar()));
  $("#esvaziar")?.addEventListener("click", () => {
    if (!confirm("Tirar todos os modelos do carrinho?")) return;
    carrinho.esvaziar();
    redesenharCarrinho();
  });
  $("#encomendar")?.addEventListener("click", prepararEncomenda);
  $(".cart__lista", app)?.addEventListener("click", (e) => {
    const item = e.target.closest(".citem");
    if (!item) return;
    const p = produtoPorRef(item.dataset.ref);
    if (e.target.closest("[data-editar]")) abrirEscolha(p);
    if (e.target.closest("[data-remover]")) {
      carrinho.definir(p.ref, {});
      redesenharCarrinho();
      aviso(`<span>${esc(p.codigo)} saiu do carrinho</span>`);
    }
  });
}

/* ---------- enviar a encomenda ---------- */
function carregarImagem(src) {
  return new Promise((ok) => {
    const im = new Image();
    im.onload = () => ok(im);
    im.onerror = () => ok(null);
    im.src = src;
  });
}

// primeira foto que exista: cores encomendadas primeiro, depois as outras
async function fotoModelo(p, preferidas) {
  const ordem = [...preferidas, ...p.cores.filter((c) => !preferidas.includes(c))];
  for (const cor of ordem)
    for (const f of ["png", "mini", "PNG"]) {
      const im = await carregarImagem(caminhoFoto(p.ref, cor, f));
      if (im) return im;
    }
  return null;
}

function quebrarTexto(ctx, texto, largura) {
  const palavras = texto.split(" ");
  const linhas = [];
  let atual = "";
  for (const w of palavras) {
    const tentativa = atual ? `${atual} ${w}` : w;
    if (ctx.measureText(tentativa).width > largura && atual) (linhas.push(atual), (atual = w));
    else atual = tentativa;
  }
  if (atual) linhas.push(atual);
  return linhas;
}

function retanguloRedondo(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

// Imagem (PNG) da encomenda: foto de cada modelo com as cores e números à frente
async function imagemEncomenda() {
  const fontes = ['400 40px "Instrument Serif"', 'italic 400 40px "Instrument Serif"', "500 20px Manrope", "600 20px Manrope"];
  await Promise.all(fontes.map((f) => document.fonts?.load(f).catch(() => {})));
  const refs = Object.keys(carrinho.itens);
  const fotos = await Promise.all(refs.map((ref) => fotoModelo(produtoPorRef(ref), Object.keys(carrinho.itens[ref]))));

  const COR = { fundo: "#fffef7", areia: "#f6f0e3", escuro: "#664b32", castanho: "#8f6946", tinta: "#2f2318", suave: "#8a7766", linha: "rgba(102,75,50,0.18)" };
  const SERIF = '"Instrument Serif", "Times New Roman", serif';
  const SANS = 'Manrope, system-ui, sans-serif';
  const W = 1080, M = 64, FOTO = 220, X = M + FOTO + 40, LARG = W - M - X;

  const c = document.createElement("canvas");
  const ctx = c.getContext("2d");

  // 1.º passo: medir cada modelo
  const blocos = refs.map((ref, i) => {
    const p = produtoPorRef(ref);
    ctx.font = `500 23px ${SANS}`;
    const linhas = Object.entries(carrinho.itens[ref]).map(([cor, l]) => ({
      cor,
      pares: paresLinha(p, l),
      desc: quebrarTexto(ctx, descreverLinha(p, l), LARG - 34),
    }));
    const altTexto = 60 + 40 + linhas.reduce((a, l) => a + 38 + l.desc.length * 32 + 14, 0);
    return { p, foto: fotos[i], linhas, alt: Math.max(FOTO, altTexto) + 56 };
  });
  ctx.font = `500 24px ${SANS}`;
  const notas = carrinho.notas.trim() ? quebrarTexto(ctx, carrinho.notas.trim(), W - 2 * M) : [];
  const cliente = carrinho.nome.trim();
  const ALT_TOPO = cliente ? 296 : 250;
  const ALT_FIM = 190 + (notas.length ? 50 + notas.length * 34 : 0);
  c.width = W;
  c.height = ALT_TOPO + blocos.reduce((a, b) => a + b.alt, 0) + ALT_FIM;

  // 2.º passo: desenhar
  ctx.fillStyle = COR.fundo;
  ctx.fillRect(0, 0, W, c.height);
  ctx.textBaseline = "alphabetic";

  // topo: logótipo, data e cliente
  ctx.fillStyle = COR.escuro;
  ctx.beginPath();
  ctx.arc(M + 32, M + 32, 32, 0, Math.PI * 2);
  ctx.fill();
  ctx.fillStyle = COR.fundo;
  ctx.font = `400 28px ${SERIF}`;
  ctx.textAlign = "center";
  ctx.fillText("MV", M + 32, M + 42);
  ctx.textAlign = "left";
  ctx.font = `400 46px ${SERIF}`;
  ctx.fillStyle = COR.escuro;
  ctx.fillText("Maneira", M + 84, M + 48);
  const wM = ctx.measureText("Maneira ").width;
  ctx.font = `italic 400 46px ${SERIF}`;
  ctx.fillStyle = COR.castanho;
  ctx.fillText("Vaidosa", M + 84 + wM, M + 48);

  ctx.textAlign = "right";
  ctx.font = `600 18px ${SANS}`;
  ctx.fillStyle = COR.castanho;
  ctx.fillText("ENCOMENDA", W - M, M + 22);
  ctx.font = `500 22px ${SANS}`;
  ctx.fillStyle = COR.suave;
  ctx.fillText(new Date().toLocaleDateString("pt-PT", { day: "numeric", month: "long", year: "numeric" }), W - M, M + 52);
  ctx.textAlign = "left";

  ctx.font = `400 64px ${SERIF}`;
  ctx.fillStyle = COR.tinta;
  ctx.fillText(cliente || CONFIG.colecao, M, M + 160);
  if (cliente) {
    ctx.font = `500 22px ${SANS}`;
    ctx.fillStyle = COR.suave;
    ctx.fillText(CONFIG.colecao, M, M + 196);
  }
  let y = ALT_TOPO;
  ctx.fillStyle = COR.linha;
  ctx.fillRect(M, y - 1, W - 2 * M, 1);

  for (const b of blocos) {
    const y0 = y + 28;
    // foto
    ctx.fillStyle = COR.areia;
    retanguloRedondo(ctx, M, y0, FOTO, FOTO, 28);
    ctx.fill();
    ctx.fillStyle = "rgba(255,254,247,0.75)";
    ctx.beginPath();
    ctx.arc(M + FOTO / 2, y0 + FOTO / 2, FOTO * 0.36, 0, Math.PI * 2);
    ctx.fill();
    if (b.foto) {
      const esc_ = Math.min((FOTO * 0.86) / b.foto.width, (FOTO * 0.72) / b.foto.height);
      const w = b.foto.width * esc_, h = b.foto.height * esc_;
      ctx.drawImage(b.foto, M + (FOTO - w) / 2, y0 + (FOTO - h) / 2, w, h);
    }
    // referência, categoria, preço
    let ty = y0 + 46;
    ctx.font = `400 54px ${SERIF}`;
    ctx.fillStyle = COR.escuro;
    ctx.fillText(b.p.codigo, X, ty);
    ctx.textAlign = "right";
    ctx.font = `600 20px ${SANS}`;
    ctx.fillStyle = COR.castanho;
    ctx.fillText(b.p.precoCodigo, W - M, ty);
    ctx.textAlign = "left";
    ty += 40;
    ctx.font = `600 17px ${SANS}`;
    ctx.fillStyle = COR.suave;
    ctx.fillText(b.p.categoria.toUpperCase(), X, ty);
    ty += 22;
    // cores
    for (const l of b.linhas) {
      ty += 36;
      const sw = corSwatch(l.cor);
      const hex = sw.match(/#[0-9a-f]{3,6}/gi);
      if (sw.startsWith("linear") && hex) {
        const g = ctx.createLinearGradient(X, ty - 24, X + 22, ty - 2);
        hex.forEach((h, k) => g.addColorStop(k / (hex.length - 1), h));
        ctx.fillStyle = g;
      } else ctx.fillStyle = sw;
      ctx.beginPath();
      ctx.arc(X + 11, ty - 8, 11, 0, Math.PI * 2);
      ctx.fill();
      ctx.strokeStyle = "rgba(102,75,50,0.35)";
      ctx.lineWidth = 1.5;
      ctx.stroke();
      ctx.font = `600 25px ${SANS}`;
      ctx.fillStyle = COR.tinta;
      ctx.fillText(l.cor, X + 34, ty);
      ctx.textAlign = "right";
      ctx.font = `600 21px ${SANS}`;
      ctx.fillStyle = COR.castanho;
      ctx.fillText(plur(l.pares, "par", "pares"), W - M, ty);
      ctx.textAlign = "left";
      ctx.font = `500 23px ${SANS}`;
      ctx.fillStyle = COR.suave;
      for (const linha of l.desc) {
        ty += 32;
        ctx.fillText(linha, X + 34, ty);
      }
      ty += 14;
    }
    y += b.alt;
    ctx.fillStyle = COR.linha;
    ctx.fillRect(M, y - 1, W - 2 * M, 1);
  }

  // totais e notas
  const t = carrinho.totais();
  y += 64;
  ctx.font = `600 18px ${SANS}`;
  ctx.fillStyle = COR.castanho;
  ctx.fillText("TOTAL", M, y);
  y += 62;
  ctx.font = `400 60px ${SERIF}`;
  ctx.fillStyle = COR.escuro;
  ctx.fillText(`${plur(t.modelos, "modelo", "modelos")} · ${plur(t.pares, "par", "pares")}`, M, y);
  if (notas.length) {
    y += 56;
    ctx.font = `600 18px ${SANS}`;
    ctx.fillStyle = COR.castanho;
    ctx.fillText("NOTAS", M, y);
    ctx.font = `500 24px ${SANS}`;
    ctx.fillStyle = COR.tinta;
    for (const linha of notas) (y += 34), ctx.fillText(linha, M, y);
  }

  return new Promise((ok) => c.toBlob(ok, "image/png"));
}

async function prepararEncomenda() {
  const btn = $("#encomendar");
  const txt = $("span", btn);
  btn.disabled = true;
  txt.textContent = "A preparar a encomenda…";
  let blob = null;
  try {
    blob = await imagemEncomenda();
  } catch (err) {
    console.error(err);
  }
  btn.disabled = false;
  txt.textContent = "Enviar encomenda";

  const data = new Date().toISOString().slice(0, 10);
  const nomeFicheiro = `encomenda-maneira-vaidosa-${data}.png`;
  const ficheiro = blob ? new File([blob], nomeFicheiro, { type: "image/png" }) : null;
  const url = blob ? URL.createObjectURL(blob) : "";
  const podePartilhar = !!(ficheiro && navigator.canShare?.({ files: [ficheiro] }));

  if (!blob) {
    alert("Não foi possível preparar a imagem da encomenda. Tenta outra vez.");
    return;
  }
  const painel = abrirPainel(
    `<header class="painel__topo">
      <div><p class="eyebrow">Enviar encomenda</p><h2 class="painel__titulo">Quase feito</h2></div>
      <button type="button" class="painel__x" data-fechar aria-label="Fechar">×</button>
    </header>
    <div class="painel__corpo">
      <div class="enc__preview"><img src="${url}" alt="Imagem da encomenda"></div>
      <p class="enc__ajuda">${
        podePartilhar
          ? "Carrega em <b>Enviar</b>, escolhe o <b>WhatsApp</b> e a conversa da Maneira Vaidosa."
          : "Descarrega a imagem e envia-a pelo WhatsApp para a Maneira Vaidosa."
      }</p>
      <div class="enc__acoes">
        ${podePartilhar ? `<button type="button" class="btn" id="encEnviar">${ICONE_ENVIAR}Enviar</button>` : ""}
        <button type="button" class="btn ${podePartilhar ? "btn--ghost" : ""}" id="encDescarregar">Descarregar imagem</button>
      </div>
      <p class="enc__nota" id="encNota" hidden>Encomenda enviada? <button type="button" class="link-btn" id="encLimpar">Esvaziar o carrinho</button></p>
    </div>`,
    "painel--encomenda"
  );
  const mostrarNota = () => ($("#encNota", painel).hidden = false);

  $("#encEnviar", painel)?.addEventListener("click", async () => {
    try {
      await navigator.share({ files: [ficheiro] });
      mostrarNota();
    } catch {}
  });
  $("#encDescarregar", painel).addEventListener("click", () => {
    const a = document.createElement("a");
    a.href = url;
    a.download = nomeFicheiro;
    document.body.appendChild(a);
    a.click();
    a.remove();
    mostrarNota();
  });
  $("#encLimpar", painel).addEventListener("click", () => {
    carrinho.esvaziar();
    fecharPainel();
    redesenharCarrinho();
    aviso("<span>Obrigada pela encomenda!</span>");
  });
}

/* ---------- router ---------- */
function rota() {
  const h = decodeURIComponent(location.hash.replace(/^#\/?/, ""));
  const [tipo, val] = h.split("/");
  if (tipo === "c" && val) return { nome: "categoria", val, nav: val };
  if (tipo === "p" && val) return { nome: "produto", val, nav: estado.produtos.find((p) => p.ref === val)?.catSlug };
  if (tipo === "contactos") return { nome: "contactos", nav: "contactos" };
  if (tipo === "carrinho") return { nome: "carrinho", nav: "carrinho" };
  if (tipo === "" || tipo === "colecao") return { nome: "inicio", nav: "todos" };
  return { nome: "404" };
}

let ultimaRota = null;
function render() {
  // âncoras internas (#colecao) não mudam de página
  if (location.hash === "#colecao") return;
  const r = rota();
  const app = $("#app");

  $("#nav").innerHTML = htmlNav(r.nav);
  document.body.dataset.page = r.nome;

  app.classList.remove("is-entering");
  void app.offsetWidth;
  app.innerHTML =
    r.nome === "inicio" ? paginaInicio() :
    r.nome === "categoria" ? paginaCategoria(r.val) :
    r.nome === "produto" ? paginaProduto(r.val) :
    r.nome === "contactos" ? paginaContactos() :
    r.nome === "carrinho" ? paginaCarrinho() : pagina404();
  app.classList.add("is-entering");

  // ao voltar de uma categoria para "Todos", salta diretamente para a grelha
  const colecao = $("#colecao");
  const top = ultimaRota?.nome === "categoria" && r.nome === "inicio" && colecao ? colecao.offsetTop - 60 : 0;
  window.scrollTo({ top, behavior: "instant" });

  const titulos = {
    inicio: `${CONFIG.empresa.nome} — ${CONFIG.colecao}`,
    categoria: `${plural(estado.categorias.find((c) => c.slug === r.val)?.nome || "")} — ${CONFIG.empresa.nome}`,
    produto: `${estado.produtos.find((p) => p.ref === r.val)?.codigo || ""} — ${CONFIG.empresa.nome}`,
    contactos: `Contactos — ${CONFIG.empresa.nome}`,
    carrinho: `Carrinho — ${CONFIG.empresa.nome}`,
  };
  document.title = titulos[r.nome] || CONFIG.empresa.nome;

  ativarReveal();
  ativarCards();
  if (r.nome === "produto") ativarProduto(r.val);
  if (r.nome === "inicio") ativarHero();
  if (r.nome === "carrinho") ativarCarrinho();
  $("#cartBtn")?.classList.toggle("is-active", r.nome === "carrinho");
  fecharPainel();
  if (r.nome === "carrinho") $(".toast")?.remove();
  // centra o link ativo na barra de filtros (telemóvel)
  $("#nav .is-active")?.scrollIntoView({ block: "nearest", inline: "center" });
  ultimaRota = r;
}

/* ---------- arranque ---------- */
document.addEventListener("click", (e) => {
  const a = e.target.closest("[data-scroll]");
  if (!a) return;
  e.preventDefault();
  $("#colecao")?.scrollIntoView({ behavior: "smooth", block: "start" });
});

const topbar = $("#topbar");
const aoScroll = () => topbar.classList.toggle("is-scrolled", window.scrollY > 24);
window.addEventListener("scroll", aoScroll, { passive: true });

$("#year").textContent = new Date().getFullYear();

(async function iniciar() {
  try {
    await carregarDados();
    carrinho.carregar();
    atualizarContador();
    window.addEventListener("hashchange", render);
    render();
    aoScroll();
  } catch (err) {
    console.error(err);
    $("#app").innerHTML = `<section class="page-head page-head--center">
      <p class="eyebrow">Erro</p>
      <h1 class="page-title">Não foi possível carregar o catálogo</h1>
      <p class="muted">${esc(err.message)}.<br>Se abriu o ficheiro diretamente, o site precisa de ser servido por um servidor (ver README).</p>
    </section>`;
  }
})();
