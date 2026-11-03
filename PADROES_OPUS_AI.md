# Padrões de engenharia — Opus AI

Documento de referência para **todos** os projetos da Opus AI. Antes de começar
qualquer plataforma nova, leia este arquivo. Antes de discutir uma decisão que
parece familiar, procure aqui primeiro — provavelmente já foi resolvida.

Última revisão: 27/08/2026 · originado do projeto COB.

---

## 1. Arquitetura de arquivos (Google Apps Script)

Nunca deixe o frontend num único arquivo. Um `index.html` de 2.000 linhas é
impossível de manter e inviabiliza o uso de assistentes de código — cada leitura
do arquivo consome a janela de contexto inteira.

Divisão padrão:

| Arquivo | Conteúdo | Pode ter `<? ?>` |
|---|---|---|
| `Code.gs` | Todo o backend | — |
| `index.html` | Roteamento, valores do servidor, includes | **Sim** |
| `Estilo.html` | Todo o CSS | Não |
| `Formulario.html` | Marcação das telas públicas | Não |
| `Painel.html` | Marcação do painel administrativo | Não |
| `Textos.html` | Dicionário de traduções (i18n) | Não |
| `Script.html` | Toda a lógica JavaScript | Não |

No `Code.gs`:

```js
function include(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}
```

No `index.html`, os valores do servidor ficam num bloco próprio, **antes** dos
includes de JavaScript:

```html
<script>
var APP_URL = "<?= appUrl ?>";
var VIEW_MODE = "<?= viewMode ?>";
</script>
<?!= include('Textos') ?>
<?!= include('Script') ?>
```

**Regra que evita a maior fonte de tela branca:** arquivos incluídos não podem
conter scriptlets. `include()` não avalia template — todo `<? ?>` mora no
`index.html`. A ordem importa: `Textos` antes de `Script`.

---

## 2. Dados: fonte única da verdade

**A planilha é a fonte da verdade. Google Calendar, Gmail e qualquer outro
serviço são projeções dela.**

- Nenhuma escrita no Calendar acontece antes de a planilha ter sido gravada.
- Toda alteração de um registro passa por **uma única função** no backend.
  Ordem obrigatória: validar sessão → `LockService` → validar dados → gravar
  planilha → atualizar Calendar → enviar e-mail → devolver resultado.
- Ao alterar um evento existente, use `setTime()`. **Nunca** apague e recrie:
  recriar faz o convidado perder o RSVP e gera convite duplicado.
- Toda plataforma precisa de um **reconciliador** que compare planilha e Calendar
  e relate divergências. Ele relata; quem corrige é o administrador.
- Registre histórico das alterações numa coluna própria:
  `dd/MM/yyyy HH:mm · AÇÃO · de X para Y · por quem`.

---

## 3. Diagramação de e-mail

Simetria é requisito, não estética.

- **Corpo:** alinhado à esquerda.
- **Notas de rodapé, assinaturas e avisos:** centralizados, em coluna estreita
  (`max-width` entre 340px e 440px), com quebras definidas manualmente (`<br>`)
  para que as duas linhas tenham extensão semelhante.
- **Nunca deixe palavra órfã sozinha na segunda linha.** Se sobrar, reescreva a
  frase até equilibrar. Cabeçalhos com `letter-spacing` alto precisam de
  `max-width` folgado, senão quebram feio.
- **Comprovante:** todo e-mail de confirmação repete integralmente os dados que a
  pessoa forneceu. Serve como documentação para ela.
- **Rodapé de LGPD obrigatório** em todo e-mail dirigido ao público: citar a Lei
  nº 13.709/2018, declarar que se coleta apenas o necessário e não se compartilha
  com terceiros, e informar o canal para acesso, correção e exclusão de dados.
- Renderize o e-mail antes de entregar. Não confie no código; olhe o resultado.

---

## 4. Interface

- **Mobile-first.** Testar sempre a **375px e 320px**. Zero overflow horizontal.
- Alvos de toque de no mínimo **56px**.
- Prefira botões, chips e seletores a digitação livre. Menos toques, menos erro.
- Uma pergunta por tela em formulários longos.
- Telas condicionais (`showIf`) em vez de perguntas irrelevantes.
- Confirmações e diálogos usam **modal próprio da plataforma**.
  `confirm()` e `alert()` são proibidos.
- Na tela de revisão final, cada dado tem botão de editar que leva à tela certa e
  volta para a revisão. **Calcule o índice da tela em tempo de execução** — as
  telas visíveis mudam conforme as respostas; índice fixo manda para o lugar errado.

### Tabelas

- `table-layout: fixed` com larguras percentuais explícitas por coluna.
- `font-variant-numeric: tabular-nums` em datas e números.
- Selos (`badge`) com `min-width` igual, para todos terem a mesma caixa.
- Botões de ação com `min-width` e `height` iguais, sempre na mesma ordem.
  Quando uma ação não existe naquela linha, renderize um espaçador invisível do
  mesmo tamanho — as colunas de botões precisam alinhar entre linhas.
- Ritmo vertical constante: mesmo `padding` e mesma `line-height` em todo `td`.
  Se um campo faltar, use `—` para não encolher a linha.
- No modo cartão (≤720px), o rótulo tem largura fixa para os valores alinharem.

### Acessibilidade e clareza

- `title` em todo botão de ação, com frase explicando o que ele faz.
- `title` não existe em telas de toque: acrescente também uma **legenda fixa**
  visível abaixo da tabela.
- `aria-label` em todo botão que só tem ícone.

---

## 5. Segurança

- Toda função que lê ou grava dado pessoal exige **sessão validada no servidor**.
  Validação só no navegador não é validação — qualquer pessoa com a URL chama a
  função direto.
- Links de ação enviados por e-mail são **assinados com HMAC**. ID adivinhável
  não é credencial.
- Limite de tentativas no código de acesso e intervalo mínimo entre pedidos.
- Todo dado vindo de formulário público é **escapado** antes de ir para
  `innerHTML`. Nome de candidato é entrada não confiável.
- Valide no servidor tudo que você validou no navegador. O cliente mente.

---

## 6. Google Apps Script — armadilhas conhecidas

- **Implantação:** sempre `Implantar → Gerenciar → lápis → Nova versão`.
  **Nunca** "Nova implantação" — ela troca a URL e quebra todos os links já
  distribuídos e os `_redirects`.
- Rodar a função de diagnóstico **antes** de cada publicação, conferindo em
  especial a autorização de escrita no Calendar.
- `localStorage` e `sessionStorage` lançam `SecurityException` dentro do iframe.
  Proibidos. Use estado em memória.
- Funções `.gs` duplicadas entre arquivos causam erro de compilação silencioso:
  tela branca, nenhuma mensagem.
- `CacheService`: TTL máximo de 21.600 segundos.
- `LockService`: `waitLock(20000)` com `releaseLock()` em bloco `finally`.
- Telefone com `+` precisa do truque do apóstrofo (`"'" + numero`), senão o Sheets
  interpreta como fórmula e grava `#ERROR`.
- **Fuso horário:** `getDisplayValues()` renderiza datas no fuso **da planilha**,
  não no do script. Force `ss.setSpreadsheetTimeZone("America/Sao_Paulo")` e leia
  datas de `getValues()`, formatando no servidor com `Utilities.formatDate`.
- Uma única chamada `CalendarApp.getEvents(inicio, fim)` para toda a janela.
  Uma chamada por horário estoura o limite de 6 minutos de execução.
- Nunca deixe fallback silencioso em recurso externo. Se a agenda configurada
  falhar, **lance erro visível** — cair na agenda padrão sem avisar é pior do que
  quebrar.
- Erro em fonte de dados devolve objeto de erro, nunca lista vazia: lista vazia
  vira "não há horários", que é mentira.

---

## 7. Verificação antes de entregar

Não relate teste que não foi executado.

```bash
cp Code.gs /tmp/c.js && node --check /tmp/c.js          # sintaxe do backend
node --check /tmp/app.js                                 # sintaxe do JS extraído
grep -c '<?' Arquivo.html                                # scriptlets em incluído = 0
grep -n "localStorage\|sessionStorage" *.html            # precisa ser vazio
grep -n "confirm(\|alert(" *.html                        # precisa ser vazio
```

Mais: fluxo completo no Playwright a 375px e 320px, zero overflow horizontal,
zero erro de JavaScript no console, e inspeção visual dos e-mails renderizados.

---

## 8. Comercial

- Nunca propor ferramenta antes de completar o diagnóstico de processo.
- Nunca cotar preço na primeira reunião.
- Diferencial do lembrete por e-mail (24h e 2h antes, para cliente e profissional,
  independente do aceite no Google Calendar) entra em toda proposta.
- Astrea integra nativamente com Google Drive — recomendar Drive, não OneDrive,
  para clientes jurídicos que usam Astrea.
- Mensagens de WhatsApp só como link `wa.me` pré-preenchido. Nunca envio automático.
