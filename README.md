# Plataforma COB — Conectando Orquestras Brasileiras

Cadastro e onboarding de **voluntários, embaixadores e representantes de orquestra**.
A pessoa se inscreve, sugere três horários, e a coordenação confirma um deles — o que
cria o evento no Google Calendar e dispara os e-mails automaticamente.

Construída em Google Apps Script sobre Google Sheets, Calendar, Drive e Gmail.
Desenvolvimento: **Opus AI / Studio Kephra**.

---

## Arquivos canônicos

Apenas estes oito arquivos são código. Qualquer outra coisa na raiz é resíduo e
deve ser apagada.

| Arquivo | O que é | Aceita `<? ?>` |
|---|---|---|
| `Code.gs` | Backend completo | — |
| `appsscript.json` | Manifesto: fuso, escopos, publicação | — |
| `index.html` | Roteamento, valores do servidor e includes | **Sim** |
| `Estilo.html` | CSS | **Não** |
| `Formulario.html` | Marcação das telas públicas | **Não** |
| `Painel.html` | Marcação do painel administrativo | **Não** |
| `Textos.html` | Dicionário de traduções PT / EN / ES | **Não** |
| `Script.html` | Toda a lógica de frontend | **Não** |

Documentação: `README.md` e `PADROES_OPUS_AI.md`.

---

## Arquitetura do frontend

O frontend foi dividido em cinco arquivos porque um `index.html` de duas mil linhas
é impossível de manter — e inviabiliza o uso de assistentes de código, que gastam a
janela de contexto inteira só para ler o arquivo.

A montagem acontece em `index.html`:

```html
<?!= include('Estilo') ?>
...
<?!= include('Formulario') ?>
<?!= include('Painel') ?>
<script>
var APP_URL = "<?= appUrl ?>";
var VERSAO  = "<?= versao ?>";
</script>
<?!= include('Textos') ?>
<?!= include('Script') ?>
```

### Duas regras que causam tela branca sem mensagem de erro

**1. Arquivos incluídos não podem conter scriptlets.** `include()` usa
`createHtmlOutputFromFile`, que não avalia template. Todo valor vindo do servidor
mora no `index.html`; se `Script.html` precisar de um dado do backend, declare a
variável no `index.html` e leia de lá.

**2. `Textos` vem antes de `Script`.** A ordem dos includes não é cosmética.

Antes de subir qualquer alteração, confirme:

```bash
grep -c '<?' Estilo.html Formulario.html Painel.html Textos.html Script.html
# precisa devolver 0 em todos
```

---

## Como publicar

1. No editor do Apps Script, substitua o conteúdo dos oito arquivos.
   Ao criar um arquivo HTML novo, digite o nome **sem** a extensão — o GAS
   acrescenta o `.html`. O nome tem que bater exatamente com o do `include()`.
2. Rode **`verificarInstalacao`** no editor e leia o log. Todas as linhas devem
   começar com `OK`, em especial `Calendar (escrita)`.
3. **Implantar → Gerenciar implantações → ✏️ (lápis) → Versão: Nova versão → Implantar.**

> ⚠️ **Nunca use "Nova implantação".** Ela gera uma URL diferente e quebra todos os
> links já distribuídos por e-mail e o botão no site do COB.

### Confirmando qual versão está no ar

O `Code.gs` carrega uma constante `VERSAO`. Para saber em segundos se o deploy pegou:

- **Console do navegador** (F12) → `Plataforma COB — versão ...`
- **Painel administrativo** → a faixa de diagnóstico começa com `v ...`
- **Editor** → `verificarInstalacao` abre com `VERSÃO PUBLICADA: ...`

Se o console não mostrar a versão esperada, o problema é a etapa 3 acima, não o código.

---

## Fonte única da verdade

**A planilha manda. Google Calendar e Gmail são projeções dela.**

- Nenhuma escrita no Calendar acontece antes de a planilha ter sido gravada.
- Toda alteração passa por uma função só, na ordem: validar sessão → `LockService`
  → validar dados → gravar planilha → atualizar Calendar → enviar e-mail.
- Eventos existentes são **editados**, nunca apagados e recriados: recriar faz o
  convidado perder o RSVP.
- `_reconciliarAgenda` compara planilha e Calendar e **relata** divergências.
  Quem corrige é o administrador, pelo painel.

---

## Abas da planilha

| Aba | Função |
|---|---|
| `Voluntarios`, `Embaixadores`, `Representantes` | Cadastros, uma linha por pessoa |
| `Admins` | E-mails autorizados. Quem está aqui recebe as notificações **e** acessa o painel |
| `Configuracoes_Agenda` | Janela de horários oferecidos |
| `Estrutura_COB` | Estrutura administrativa (módulo desativado nesta fase) |

`Configuracoes_Agenda`, linha 2 — **números inteiros, nunca `10:00`**:

| HORA_INICIO | HORA_FIM | DIAS_HORIZONTE | ANTECEDENCIA_HORAS | INCLUIR_SABADO |
|---|---|---|---|---|
| 10 | 18 | 21 | 24 | NAO |

Se a aba estiver no formato antigo, `garantirAbasBase` a normaliza sozinha.

Não altere a ordem nem os nomes de `HEADERS_BASE` e do objeto `C` no `Code.gs`.
Colunas novas só entram **no fim**.

---

## Funções de manutenção

Rodar no editor do Apps Script, selecionando a função e clicando em *Executar*.

| Função | Para quê |
|---|---|
| `verificarInstalacao()` | Diagnóstico geral: escopos, agenda, pasta, cota de e-mail, versão |
| `diagnosticarPainel()` | Quando um cadastro não aparece no painel. Mostra o que existe em cada aba e o que o painel leria |
| `TESTE_agenda()` | Confirma que a agenda configurada é legível **e** gravável |
| `TESTE_email_de_notificacao()` | Confirma entrega de e-mail |
| `usarAgendaDoCOB()` | Aponta a plataforma para a agenda do COB |
| `compartilharPastaRaiz()` | Dá acesso de editor à pasta do Drive |
| `instalarGatilhoLembretes()` | Instala o gatilho diário de lembretes |
| `configurarProprietario({...})` | Troca pasta, planilha, agenda e contato — usado na transferência de posse |
| `reverterProprietario()` | Desfaz o anterior e volta aos padrões |

---

## Segurança

- Toda função que lê ou grava dado pessoal exige sessão validada **no servidor**.
  Validação só no navegador não é validação.
- Links de ação enviados por e-mail são assinados com HMAC.
- Limite de tentativas no código de acesso e intervalo mínimo entre pedidos.
- Todo dado vindo de formulário público é escapado antes de ir para `innerHTML`.

Os dados coletados incluem informações sensíveis (cor/raça, gênero, data de
nascimento). O tratamento segue a Lei nº 13.709/2018 e todo e-mail dirigido ao
público traz o rodapé de LGPD com o canal para acesso, correção e exclusão.

---

## Antes de abrir um pull request

```bash
cp Code.gs /tmp/c.js && node --check /tmp/c.js       # sintaxe do backend
grep -c '<?' Script.html Textos.html                  # precisa ser 0
grep -n "localStorage\|sessionStorage" *.html         # precisa ser vazio
grep -n "confirm(\|alert(" *.html                     # precisa ser vazio
```

E teste o fluxo completo a **375px e 320px**, com zero overflow horizontal.
As convenções que valem para todos os projetos estão em `PADROES_OPUS_AI.md`.
