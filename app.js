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
  empresa: {
    nome: "Maneira Vaidosa",
    telefone: "256 812 118",
    morada: ["Travessa do Emigrante, Nº 120", "3700-475 Arrifana, Portugal"],
    // Número que recebe as encomendas por WhatsApp (só é usado no botão "Encomendar", não aparece no site)
    whatsappEncomendas: "351913452242",
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

// Mensagem de encomenda por WhatsApp com a referência, a cor e o link da foto do modelo
function linkWhatsApp(p, cor) {
  const foto = new URL(caminhoFoto(p.ref, cor), location.href).href;
  const texto =
    `Olá! Gostaria de encomendar o modelo ${p.codigo} (${p.categoria}), cor ${cor}.\n` +
    `Foto: ${foto}`;
  return `https://wa.me/${CONFIG.empresa.whatsappEncomendas}?text=${encodeURIComponent(texto)}`;
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
  return semAcentos(h).toLowerCase().replace(/[^a-z]/g, "");
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
      };
      mapa.set(l.ref, p);
    }
    if (l.cor && !p.cores.includes(l.cor)) p.cores.push(l.cor);
  }
  estado.produtos = [...mapa.values()].filter((p) => p.cores.length);
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

  const tamanhos = (() => {
    const m = p.numeracao.match(/^(\d+)\s*[-–a]\s*(\d+)$/);
    if (!m) return `<span class="size">${esc(p.numeracao || "—")}</span>`;
    const out = [];
    for (let n = +m[1]; n <= +m[2]; n++) out.push(`<span class="size">${n}</span>`);
    return out.join("");
  })();

  return `
  <section class="product">
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
          <a class="btn btn--wa" id="encomendar" href="${linkWhatsApp(p, p.corInicial)}" target="_blank" rel="noopener">
            <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path fill="currentColor" d="M12 2a10 10 0 0 0-8.6 15.1L2 22l5-1.3A10 10 0 1 0 12 2Zm0 18.2a8.2 8.2 0 0 1-4.2-1.2l-.3-.2-3 .8.8-2.9-.2-.3A8.2 8.2 0 1 1 12 20.2Zm4.5-6.1c-.2-.1-1.5-.7-1.7-.8-.2-.1-.4-.1-.6.1l-.8 1c-.1.2-.3.2-.5.1a6.7 6.7 0 0 1-3.3-2.9c-.3-.4.2-.4.7-1.4.1-.2 0-.3 0-.4l-.8-1.8c-.2-.5-.4-.4-.6-.4h-.5a1 1 0 0 0-.7.3 3 3 0 0 0-.9 2.2 5.2 5.2 0 0 0 1.1 2.7 11.8 11.8 0 0 0 4.5 4c1.7.7 2.3.8 3.2.6.5-.1 1.5-.6 1.7-1.2.2-.6.2-1.1.2-1.2-.1-.1-.3-.2-.5-.3Z"/></svg>
            Encomendar
          </a>
          <a class="btn btn--ghost" href="#/contactos">Contactos</a>
        </div>
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
    $("#encomendar").href = linkWhatsApp(p, cor);
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

/* ---------- router ---------- */
function rota() {
  const h = decodeURIComponent(location.hash.replace(/^#\/?/, ""));
  const [tipo, val] = h.split("/");
  if (tipo === "c" && val) return { nome: "categoria", val, nav: val };
  if (tipo === "p" && val) return { nome: "produto", val, nav: estado.produtos.find((p) => p.ref === val)?.catSlug };
  if (tipo === "contactos") return { nome: "contactos", nav: "contactos" };
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
    r.nome === "contactos" ? paginaContactos() : pagina404();
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
  };
  document.title = titulos[r.nome] || CONFIG.empresa.nome;

  ativarReveal();
  ativarCards();
  if (r.nome === "produto") ativarProduto(r.val);
  if (r.nome === "inicio") ativarHero();
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
