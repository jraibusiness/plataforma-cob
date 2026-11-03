# Brand Kit & Sistema de Design — COB (Conectando Orquestras Brasileiras)

> Diretrizes canônicas de identidade visual e componentes da **plataforma única COB**
> (site institucional + onboarding + apoio financeiro), desenvolvida pela **Opus AI / Studio Kephra**.
> Versão 2.0 — consolida o Brand Kit original e o resolve como tokens aplicáveis em código.

---

## 1. Princípios da marca

- **Elegância sinfônica contemporânea.** Solenidade e impacto artístico sem o estigma do elitismo empoeirado.
- **Acolhimento & comunidade.** Interlocução horizontal entre orquestras profissionais, jovens, comunitárias e projetos sociais de todo o Brasil.
- **Simetria & rigor estrutural (Padrão Opus AI).** Textos institucionais e formulários alinhados à esquerda; títulos, emblemas e notas de rodapé centralizados; espaçamentos em múltiplos de 4 px / 8 px.
- **Clareza acima de ornamento.** O ornamento (prata, dourado, notas musicais) é acento, nunca ruído. Contraste de leitura primeiro.

---

## 2. Paleta cromática canônica (tokens)

Base **grafite** (não verde-chapado), verde institucional como **acento de identidade**, prata como **ação primária**, dourado e rosa como **detalhes de prestígio/afeto**.

| Token CSS | Hex | Aplicação |
| :-- | :-- | :-- |
| `--bg` | `#0F0F0F` | Fundo principal da página / base de contraste |
| `--bg-2` | `#1A1A1A` | Seções alternadas, superfícies neutras profundas |
| `--card` | `#242424` | Cards, modais, caixas de etapa, formulários |
| `--card-2` | `#2E2E2E` | Realce de card / hover de superfície |
| `--green` | `#2D5A3D` | Verde floresta institucional: identidade, bordas ativas, cards de perfil |
| `--green-hover` | `#1E3F2A` | Hover/foco dos elementos verdes |
| `--silver` | `#C0C0C0` | Prata nobre: **CTAs primários**, títulos Cinzel, ícones, divisores |
| `--silver-2` | `#E0E0E0` | Prata claro para textos de destaque de alta legibilidade |
| `--gold` | `#B8860B` | Dourado: marcadores de prestígio, selo de embaixador, acentos finos |
| `--pink` | `#E8A2BA` | Rosa institucional: afeto, campanhas, destaques quentes (ex.: "Seja um Amigo") |
| `--text` | `#F5F5F5` | Texto corrido principal, títulos de campo |
| `--text-mute` | `#A8A8A8` | Textos de apoio, subtítulos, instruções |
| `--border` | `rgba(192,192,192,0.14)` | Bordas sutis translúcidas em prata |
| `--danger` | `#E8736A` | Erros / ações destrutivas |
| `--success` | `#88C9A1` | Confirmações |

**Contraste (WCAG).** `--text` sobre `--bg`/`--card` ≈ 17:1 (AAA). Prata `--silver` sobre `--bg` ≈ 9:1 (AAA p/ texto grande/UI). **Nunca** usar `--text-mute` para texto essencial abaixo de 14 px. Texto escuro (`#0F0F0F`) só sobre superfícies claras (prata/gradiente de CTA).

---

## 3. Tipografia oficial

Três famílias (Google Fonts), hierarquia fixa:

1. **`Cinzel`** (serif, 600/700) — **brasão, logo, títulos de seção institucional e rótulos de CTA**. Caixa alta, `letter-spacing: 1–2px`.
2. **`Cormorant Garamond`** (serif, 500/600/700) — **hero e frases de impacto**, subtítulos humanizados e narrativos.
3. **`Inter`** (sans-serif, 400/500/600; fallback `system-ui`) — **texto corrido, formulários, painel e dados**. `line-height ≥ 1.6`.

> Resolução de inconsistência histórica: o site original usava *Playfair Display*. **Padrão único = Cinzel** para títulos/CTA; Playfair fica aposentada.

Tokens: `--font-brand: 'Cinzel'` · `--font-display: 'Cormorant Garamond'` · `--font-body: 'Inter'`.

---

## 4. Logomarcas

Arquivos oficiais (pasta `img/` do site / Drive COB):

- `COB_Circular_Branco.png` — prioritário sobre fundos escuros; favicon e avatar social.
- `COB_Horizontal_Branco.png` — topo do menu (header fixo).
- `COB_Circular_Verde.png` / `COB_Circular_Rosa.png` — campanhas, certificados, cards comemorativos.

Enquanto os PNGs não estão embutidos na plataforma (web app do Apps Script serve HTML único), usa-se um **brasão vetorial (SVG) de fallback**: círculo duplo em prata com "COB" em Cinzel. Substituir pelo PNG oficial no deploy (upload como asset/base64).

---

## 5. Componentes (UI/UX)

- **Cards de perfil (grid de 3 — Embaixador / Representante / Voluntário):** gradiente verde→grafite (`linear-gradient(135deg,#1A3D2B 0%,#151515 100%)`), borda `1px solid var(--border)`, hover com elevação de 4 px e leve brilho prateado.
- **CTA primário:** gradiente prata (`--silver`→`--silver-2`), texto `#0F0F0F`, `border-radius: 999px`, sombra suave. Hover: `translateY(-3px)` + brilho.
- **CTA secundário / contorno:** transparente, borda `1px var(--silver)`, hover iluminado.
- **Inputs:** fundo `#181818`, borda `var(--border)`; foco com contorno `--green` + halo translúcido. Alvo de toque mínimo **56 px** (mobile-first).
- **Botão de escolha (chips):** seleção por toque com auto-avanço; um por pergunta no fluxo de onboarding (padrão mobile).
- **Selo de prestígio / embaixador:** detalhe em `--gold`.

---

## 6. Diagramação & layout (Padrão Opus AI)

- Grid máx. de conteúdo: **1200 px**; leitura longa: 720–900 px.
- Ritmo vertical de seções: 72–96 px (desktop), 48–56 px (mobile).
- Gutter lateral mínimo em mobile: **16 px**; sem rolagem horizontal.
- Títulos de seção centralizados com filete prata→dourado abaixo; corpo alinhado à esquerda.
- Internacionalização: **todo texto de interface passa por i18n (PT/EN/ES)** — nada hard-coded.

---

## 7. Aplicação em campanhas

- **"Seja um Amigo do COB"** (sustentabilidade 2027): rosa `--pink` + dourado `--gold` como calor institucional sobre base grafite; planos em cards, destaque no plano Mantenedor.
- **Rede de embaixadores:** acento dourado (selo), comunicação de engajamento com régua configurável e respeito a preferências (opt-out / pausa).

---

*Mantido em `BRAND_KIT.md` na raiz do repositório `jraibusiness/plataforma-cob`. Qualquer alteração visual na plataforma deve referenciar e, se necessário, atualizar este documento.*
