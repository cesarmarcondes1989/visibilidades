# Strategic Cockpit — Context Document
> Use este arquivo para retomar o desenvolvimento em uma nova sessão. Toda a arquitetura, estado, funções, bugs conhecidos e backlog estão documentados aqui.

---

## 1. O que é

Um **arquivo HTML único** (~764 KB) que funciona como cockpit digital de estratégia para o Head of Digital Solutions & Services de uma fabricante de aeronaves (contexto Embraer). Sem servidor, sem instalação — abre direto no Chrome/Edge 86+.

**Nome do arquivo:** `COCKPIT-ABRIR-AQUI21.html` (arquivo canônico atual; versões anteriores ficam no repo só por histórico — baixe sempre o número mais alto)
**Tech stack:** Vanilla HTML + CSS + JavaScript, SheetJS (embedded, ~624 KB), File System Access API, IndexedDB
**Persistência:** Lê/escreve um arquivo `.xlsx` local via File System Access API do browser. Sem backend.
**Primeiro uso:** Cria o Excel com dados seed automaticamente.
**Usos seguintes:** Browser pede permissão para re-acessar o arquivo (comportamento normal do Chrome/Edge por sessão).

> **Convenção obrigatória de versionamento:** o usuário baixa o HTML do GitHub e abre localmente — não há deploy/refresh automático. Por isso, **toda alteração no app gera um arquivo novo com o número incrementado** (ex.: próxima mudança após `AQUI20.html` cria `AQUI21.html`), nunca editando o arquivo canônico anterior em-place. Atualizar este documento (nome do arquivo + link) e o `<title>`/referências internas se houver. Os arquivos antigos permanecem no repo como histórico.

---

## 2. Arquitetura

### Estrutura de arquivos
```
COCKPIT-ABRIR-AQUI.html   (arquivo único, ~764 KB)
  └── <script #1>: SheetJS xlsx.full.min.js (embedded, ~624 KB)
  └── <script #2>: App code (~80 KB, ~61 funções)
estrategia.xlsx            (banco de dados do usuário — Excel local)
```

### Design System
- **Tema:** Light, bg `#F3F5FA`, surface white, border `#DDE3EF`
- **Font:** Inter (UI) + IBM Plex Mono (labels, código)
- **Accent:** `#4F70F5` (blue), success `#0DAF7B`, warning `#D97706`, danger `#DC2626`
- **Shadows:** dois níveis (`--shadow`, `--shadow-md`)
- **Radius:** 16px cards, 10px botões, 999px pills/barras

---

## 3. Abas & Views

| Tab key | Label | View ID |
|---------|-------|---------|
| `visao` | Vision | `view-visao` |
| `horizontes` | Horizons | `view-horizontes` |
| `dados` | Data Strategy | `view-dados` |
| `grafo` | Graph | `view-grafo` |
| `portfolio` | Portfolio | `view-portfolio` |
| `melhorias` | Improvements & SF | `view-melhorias` |
| `budget` | 💰 Budget | `view-budget` |
| `fup` | 📋 FUP | `view-fup` |
| `admin` | ⚙ Admin | `view-admin` |

**Botão + contextual por aba:**
- Vision → `+ Theme`
- Horizons → `+ Bet`
- Data Strategy → `+ Data Product`
- Portfolio → `+ Project`
- Improvements → `+ Initiative`
- Budget → `+ Item`
- FUP → `+ Atividade`
- Graph, Admin → oculto

---

## 4. Variáveis de Estado

Todo o estado mutável vive em globals JS, persistido no Excel a cada mudança (debounce 700ms):

```javascript
let vision       = {};    // strategic intent, northStar, themes, diagPositive/Negative
let bets         = [];    // strategic bets (horizons H1/H2/H3)
let dp           = [];    // data products catalog
let dpSubs       = [];    // subprojetos de cada data product (fk dpId), tracking de implementação
let maturity     = [];    // governance maturity dimensions
let governance   = [];    // governance capability cards
let gNodes       = [];    // graph nodes (apenas não-projeto; projetos são derivados)
let gEdges       = [];    // graph edges
let projects     = [];    // portfolio projects
let mel          = [];    // improvements & SF initiatives
let dataVision   = {};    // data strategy vision statement
let principles   = [];    // data principles
let pillars      = [];    // strategic pillars (editável via Admin)
let domains      = [];    // data domains (editável via Admin)
let budget       = [];    // itens de custo (Budget tab)
let fup          = [];    // atividades de acompanhamento e informativos (FUP tab)

// Runtime only (não persistido)
let betFilter    = null;
let graphFilter  = null;
let currentTab   = "visao";
let graph        = null;  // instância ForceGraph
let _fh          = null;  // FileSystemFileHandle
let _fupTab      = "acomp"; // sub-tab ativo no FUP
let _ganttOpen   = true;  // estado expandido/recolhido do Gantt no Portfolio
let ganttDateFrom = null; // filtro de período do Gantt (string yyyy-mm-dd ou null), não persistido
let ganttDateTo   = null; // idem, data final
let ganttStatusFilter = new Set(); // filtro multi-seleção por status do Gantt, não persistido (vazio = todos)
let ganttPillarFilter = new Set(); // filtro multi-seleção por pilar do Gantt, não persistido (vazio = todos)
let _dpSubEdit   = null;  // null | "new" | id do subprojeto em edição no modal de Data Product, não persistido
```

---

## 5. Estrutura do Excel (estrategia.xlsx)

| Sheet | Conteúdo | Colunas principais |
|-------|---------|-------------|
| `_vision` | Key-value pairs | `chave`, `valor` |
| `_data_vision` | Key-value pairs | `chave`, `valor` |
| `themes` | Vision themes | `name`, `desc`, `color` |
| `principles` | Data principles | `icon`, `name`, `desc` |
| `bets` | Strategic bets | `id`, `hz`, `pillar`, `name`, `intent`, `porque`, `enablerDados`, `enablerTech`, `metrica`, `meta`, `atual`, `status` |
| `data_products` | Data products | `id`, `domain`, `name`, `owner`, `q`, `cons`, `status`, `horizon`, `sla`, `notas` |
| `dp_subs` | Subprojetos de cada data product | `id`, `dpId`, `name`, `descr`, `status`, `prog`, `start`, `end` |
| `maturity` | Governance maturity | `id`, `name`, `atual`, `alvo`, `owner` |
| `governance` | Governance capabilities | `name`, `atual`, `alvo`, `owner`, `note` |
| `projects` | Portfolio projects | `id`, `hz`, `pillar`, `bet`, `name`, `descr`, `inv`, `rec`, `prog`, `status`, `start`, `end` |
| `melhorias` | Improvements & SF | `id`, `area`, `name`, `status`, `inv`, `rec`, `prog`, `notas` |
| `graph_nodes` | Graph nodes (só não-projeto) | `id`, `type`, `label`, `group`, `x`, `y` |
| `graph_edges` | Graph edges | `from`, `to`, `intensity`, `label` |
| `pillars` | Pillars editáveis | `id`, `name`, `color` |
| `domains` | Domains editáveis | `id`, `name`, `color` |
| `budget` | Itens de custo | `id`, `name`, `categoria`, `anual`, `ytd`, `notas` |
| `fup` | Atividades FUP | `id`, `tipo`, `titulo`, `responsavel`, `prazo`, `prioridade`, `status`, `acompanhar`, `notas` |

---

## 6. Fluxo de Persistência

```
Page load
  → load()
      → _idb("get")  →  IndexedDB: recupera FileSystemFileHandle salvo
      → se encontrado & permissão "granted" → connectHandle(h)
      → se encontrado & permissão "prompt"  → showOverlay("reconnect")
      → se não encontrado                   → showOverlay("connect")

connectHandle(h)
  → h.getFile() → arrayBuffer()
  → XLSX.read(buf, {type:"array"})
  → parseWB(wb)  →  retorna {vision, bets, dp, ..., budget, fup}
  → seedAll()    →  seta todos os globals para DEFAULT_*
  → applyData(o) →  sobrescreve com dados do Excel onde não vazios
  → refreshPMAP() + refreshDMAP()
  → renderAll()
  → se Excel vazio → _write() (salva defaults)

Qualquer edição → save() → debounce 700ms → _write()
  → buildWB()  →  cria workbook de todos os globals
  → XLSX.write(wb, {bookType:"xlsx", type:"array"})
  → _fh.createWritable() → write → close
```

---

## 7. Todas as Funções do App

### Persistência & Estado
| Função | Propósito |
|----------|---------|
| `load()` | Entry point de boot — verifica IndexedDB, mostra overlay ou conecta |
| `connectHandle(h)` | Lê Excel → parseWB → applyData → renderAll |
| `_write()` | Constrói workbook e escreve no arquivo Excel |
| `save()` | Wrapper com debounce de 700ms para `_write()` |
| `seedAll()` | Reseta todo o estado para DEFAULT_* |
| `applyData(o)` | Merge dos dados do Excel nos globals |
| `buildWB()` | Cria workbook SheetJS de todo o estado |
| `parseWB(wb)` | Lê workbook SheetJS → retorna objeto de dados |
| `refreshPMAP()` | Reconstrói `PMAP` do array `pillars` |
| `refreshDMAP()` | Reconstrói `DMAP` do array `domains` |
| `_idb(mode, val)` | Helper IndexedDB: `"get"`, `"set"`, `"del"` |

### Render
| Função | Propósito |
|----------|---------|
| `renderAll()` | Chama todas as funções de render com try-catch |
| `renderVisao()` | Aba Vision |
| `renderHorizontes()` | Aba Horizons |
| `renderDados()` | Aba Data Strategy |
| `renderGrafo()` | Aba Graph |
| `renderPortfolio()` | Aba Portfolio (KPIs + Gantt + kanban por horizonte) |
| `renderGanttPortfolio()` | Desenha o Gantt de todos os projetos (agrupado por horizonte, com linha de "hoje") dentro de `#gantt-chart`, aplicando `ganttDateFrom/To` + `ganttStatusFilter` + `ganttPillarFilter` |
| `renderGanttFilters()` | Desenha os controles de filtro do Gantt (período, status, pilar) dentro de `#gantt-filters` e religa os handlers |
| `toggleGantt()` | Expande/recolhe a seção do Gantt (`_ganttOpen`), sem afetar o estado persistido |
| `buildGanttGrid(items, opts)` | Função genérica que desenha a grade de um Gantt (eixo de datas, ticks adaptativos, barras, linha de "hoje") a partir de uma lista normalizada `{id,name,status,prog,s,e,groupId?}`; usada tanto por `renderGanttPortfolio()` (com `groups:HZ`) quanto por `renderDPSubsArea()` (sem groups) |
| `renderDPSubsArea(dpId)` | Desenha dentro de `#dpSubsArea` (no modal do Data Product) o mini-Gantt dos subprojetos daquele produto + o formulário inline de criação/edição quando `_dpSubEdit` está ativo |
| `projRange(p)` | Resolve `{s,e}` (Date) efetivos de um projeto a partir de `p.start`/`p.end`, com fallback de ~3 meses quando faltam datas |
| `_pdate(s)` | Helper: converte string em `Date` válido ou `null` |
| `renderMelhorias()` | Aba Improvements & SF |
| `renderBudget()` | Aba Budget — KPIs + tabela por categoria com `<meter>` |
| `renderFup()` | Aba FUP — kanban Acompanhamento + lista Informativo |
| `renderAdmin()` | Aba Admin |
| `initGraph()` | Cria instância `ForceGraph` com nós filtrados |
| `getEffectiveGraphNodes()` | Retorna nós não-projeto do `gNodes` + nós derivados de `projects` |
| `switchTab(t)` | Troca a aba ativa |
| `switchFupTab(t)` | Troca sub-tab do FUP (acomp/info) |

### Painéis de Edição (abrem o side panel)
| Função | Abre painel para |
|----------|----------------|
| `openBet(id)` | Strategic bet (null = novo) |
| `openDP(id)` | Data product (null = novo) — **não usa o side panel**, abre o modal centralizado (`#dpModalWrap`/`openDPModal()`/`closeDPModal()`), ver "Data Strategy — Modal central do Data Product" abaixo |
| `openGov(i)` | Governance capability por index |
| `openMatPanel(i)` | Maturity dimension por index |
| `openProject(id)` | Portfolio project (null = novo) |
| `openMel(id)` | Improvement/SF initiative (null = novo) |
| `openBudget(id)` | Item de custo Budget (null = novo) |
| `openFup(id)` | Atividade FUP (null = novo) |
| `openThemePanel(i)` | Vision theme |
| `openPrinciplePanel(i)` | Data principle |
| `openDiagPanel()` | Strategic diagnosis |
| `openNodePanel(node)` | Graph node (null = novo) |
| `openAddEdgePanel(fromId)` | Add edge a partir de nó específico (multi-add com staging) |
| `openAddEdgePanelGlobal()` | Add edge do zero (multi-add com staging) |
| `openPillarPanel(id)` | Strategic pillar (null = novo) |
| `openDomainPanel(id)` | Data domain (null = novo) |

---

## 8. Comportamentos Especiais

### Graph — herança de projetos
Nós do tipo `"projeto"` no grafo são **derivados automaticamente** do array `projects` — não são armazenados separadamente. A função `getEffectiveGraphNodes()` combina nós não-projeto do `gNodes` com nós gerados dinamicamente de `projects` (ID: `"pj-" + project.id`). Posições x/y de nós de projeto são salvas em `gNodes` quando o usuário arrasta o nó.

**Importante:** qualquer código que precise exibir o `label`/`type` de um nó (incluindo nós de projeto) deve resolver via `getEffectiveGraphNodes().find(...)`, nunca via `gNodes.find(...)` diretamente — `gNodes` só contém o cache de posição (`x`/`y`) de nós de projeto já arrastados, sem `label`. Usar `gNodes.find` para resolver o nome de um nó de projeto falha silenciosamente (retorna `undefined`) e expõe o ID bruto (`pj-...`) na UI. `openNodePanel()` (lista de Conexões) e `_delEdge()` (lookup de `backToId`) foram corrigidos para usar `getEffectiveGraphNodes()`, no mesmo padrão já usado por `openAddEdgePanel()`/`openAddEdgePanelGlobal()`.

### Graph — Adicionar Conexões (multi-add)
Tanto `openAddEdgePanel()` quanto `openAddEdgePanelGlobal()` usam um sistema de **staging**: o usuário clica `+ Adicionar` várias vezes acumulando conexões numa lista visível, podendo remover com `×`, e só ao clicar `Salvar N conexões` tudo é persistido de uma vez.

### Budget — barra de execução
Usa `<meter>` HTML nativo (não `<div>`) com CSS customizado via `::-webkit-meter-*` para manter o visual consistente com o design system. Cores automáticas por threshold via atributos `low`, `high`, `optimum`. Parsing numérico defensivo via `parseFloat(String(v).replace(/[^0-9.]/g,""))` para lidar com valores do SheetJS que podem chegar como string.

### Portfolio — Gantt de todos os projetos
A aba Portfolio tem uma seção de Gantt (`renderGanttPortfolio()`) acima do board de cards, agrupada por horizonte (Now/Scale/Differentiation), com linha vertical de "hoje" e ticks de tempo adaptativos (mensal, bimestral, trimestral ou semestral conforme o intervalo total). É recolhível via `toggleGantt()` (estado `_ganttOpen`, não persistido). Clicar numa barra ou no nome do projeto (`.g-bar`/`.g-label`) abre o mesmo painel de edição (`openProject`) usado pelos cards — não há uma visão separada "somente leitura". Projetos sem `start`/`end` cadastrados recebem um intervalo padrão de ~3 meses (`projRange()`) só para exibição; os campos no Excel continuam vazios até o usuário editar.

### Portfolio — Filtros do Gantt
Acima do gráfico (`#gantt-filters`, desenhado por `renderGanttFilters()`) há três grupos de filtro, todos em estado runtime não persistido:
- **Período** (`ganttDateFrom`/`ganttDateTo`, inputs `<input type=date>` com `onchange`): quando preenchido, define o início/fim da janela visível do Gantt (em vez do range calculado automaticamente a partir dos projetos) e esconde projetos cujo intervalo não tem overlap com o período escolhido. Botão "Limpar" aparece só quando algum dos dois está preenchido.
- **Status** (`ganttStatusFilter`, `Set` com os valores de `STLIST`): chips estilo `.chip` com toggle multi-seleção (clique adiciona/remove do Set); chip "All" limpa o Set. Set vazio = sem filtro (mostra todos os status).
- **Pilar** (`ganttPillarFilter`, `Set` com `pillar.id`): mesmo padrão de chips multi-seleção, usando o array `pillars` (cores reais por pilar).
Os três filtros se combinam (AND) dentro de `renderGanttPortfolio()`, que filtra `projects` antes de calcular `ranges`/`minD`/`maxD`. Os cliques nos chips são capturados pelo listener delegado em `#view-portfolio` (`[data-gst]`/`[data-gpl]`), que chama `renderGanttPortfolio()` a cada toggle.

### Data Strategy — Modal central do Data Product
Clicar numa linha `.dp-row` (ou no botão "+ Data Product") chama `openDP(id)`, que **não usa mais o side panel** (`#scrim`/`#panel`) — em vez disso abre um modal centralizado (`#dpModalWrap` → `.cmodal`), com `openDPModal()`/`closeDPModal()` análogos a `openPanel()`/`closePanel()`. Estrutura do modal:
- `#dpModalHead`: faixa colorida (gradiente na cor do domain), nome do produto, owner, badges de status/horizonte/SLA.
- `#dpModalStats`: 3 colunas com Quality, Revenue/consumidores e Domain (omitido ao criar um produto novo).
- `#dpModalBody`: grid 2 colunas com os campos editáveis (`fld`/`sel`/`sldr`) + Notes (`ta`) + checkbox de SLA — mesma lógica de save/delete de antes.
- `#dpModalFoot`: botões Save/Delete/Close.
Fecha com o botão "✕", o botão "Close", clique fora do card (no scrim) ou Esc. Como `gv(k, root)` e `bindSliders(root)` agora aceitam um `root` opcional (default `panel`), `openDP` passa o `#dpModalBody` como root para ler/ligar os campos sem depender do side panel.

### Data Strategy — Subprojetos do Data Product
Dentro do modal do Data Product, abaixo do checkbox de SLA, há uma seção "Subprojetos" (`#dpSubsArea`, desenhada por `renderDPSubsArea(dpId)`) que permite quebrar o produto em itens menores e acompanhar o que está sendo implementado — mesma proposta do Gantt do Portfolio, só que individualizada por produto:
- Armazenados em `dpSubs` (`{id, dpId, name, descr, status, prog, start, end}`), persistidos na sheet `dp_subs` com FK `dpId`. `descr` é um campo de texto livre (textarea `ta("Descrição","sb_descr",s.descr)`) para detalhar o subprojeto. Reusam o vocabulário de status `STLIST` (Active/Bet/Exploration/Done/Paused/At Risk) e a mesma lógica de fallback de datas `projRange()` usada pelos projetos do Portfolio.
- Se já existem subprojetos, são desenhados como um mini-Gantt via `buildGanttGrid(items, {dataAttr:"sid"})` (sem agrupamento por horizonte). Clicar numa barra/label (`[data-sid]`) abre o formulário inline de edição daquele subprojeto.
- O botão "+ Subprojeto" (`#dpSubAddBtn`) seta `_dpSubEdit="new"` e mostra o formulário inline (`.dp-subform`) com nome, status, progresso (slider) e datas de início/fim — sem abrir um segundo modal sobreposto.
- Salvar/excluir um subprojeto chama `save()` imediatamente (não depende do botão "Save" do produto). `_dpSubEdit` é resetado ao salvar, cancelar, fechar o modal (`closeDPModal()`) ou ao trocar de subprojeto.
- Um Data Product **novo e ainda não salvo** não pode ter subprojetos — a área mostra "Salve o produto antes de adicionar subprojetos." e o botão "+ Subprojeto" não aparece (subprojetos dependem de um `dpId` já existente).
- Excluir um Data Product remove em cascata todos os seus `dpSubs` (`dpSubs=dpSubs.filter(s=>s.dpId!==id)` dentro do handler de `dpDlBtn`).

### FUP — dois tipos
- **`tipo: "acomp"`** — atividades de acompanhamento com Título, Responsável, Prazo, Prioridade, Status, checkbox "👁 Acompanhar". Renderizado em kanban por status.
- **`tipo: "info"`** — registros informativos com Título, Data, Notas. Renderizado em lista cronológica reversa. Preparado para uso futuro com IA.

---

## 9. Constantes Principais

### PILLARS (seed default — vira estado editável `pillars`)
```javascript
{id:"mro",   name:"Connected MRO",      color:"#0DAF7B"}
{id:"saude", name:"Health & Predictive", color:"#4F70F5"}
{id:"ops",   name:"Flight Operations",   color:"#7C3AED"}
{id:"dados", name:"Data Platform",       color:"#0891B2"}
{id:"mat",   name:"Aftermarket",         color:"#D97706"}
{id:"ia",    name:"AI & GenAI",          color:"#DB2777"}
{id:"novos", name:"New Markets",         color:"#059669"}
```

### DOMAINS (seed default — vira estado editável `domains`)
```javascript
{id:"health", name:"Aircraft Health",       color:"#0DAF7B"}
{id:"mro",    name:"MRO & Maintenance",     color:"#4F70F5"}
{id:"ops",    name:"Flight Operations",     color:"#7C3AED"}
{id:"parts",  name:"Parts & Supply",        color:"#D97706"}
{id:"cust",   name:"Customer & Commercial", color:"#DB2777"}
{id:"eng",    name:"Engineering & Quality", color:"#0891B2"}
{id:"fin",    name:"Financial",             color:"#059669"}
```

### HZ (Horizons — constante fixa)
```javascript
{id:"h1", name:"Now",            period:"0–12 months", color:"#0DAF7B"}
{id:"h2", name:"Scale",          period:"1–3 years",   color:"#4F70F5"}
{id:"h3", name:"Differentiation",period:"3–5+ years",  color:"#7C3AED"}
```

### NODE_TYPES (grafo)
```javascript
pilar:   {color:"#4F70F5", r:22, label:"Pillar"}
produto: {color:"#0DAF7B", r:17, label:"Digital Product"}
projeto: {color:"#D97706", r:13, label:"Project"}
dominio: {color:"#7C3AED", r:16, label:"Data Domain"}
```

### FUP_STATUS / FUP_PRIO
```javascript
// Status Acompanhamento
["Pendente","Em andamento","Concluído","Bloqueado","Cancelado"]

// Prioridade
["Alta","Média","Baixa"]
// Cores: Alta=#DC2626, Média=#D97706, Baixa=#0DAF7B
```

---

## 10. Padrão de Patch

```python
h = open('COCKPIT-ABRIR-AQUI.html', encoding='utf-8').read()
# ... replacements com assert para validar que o target foi encontrado ...

# Validar JS após o patch
import re, subprocess, tempfile, os
scripts = re.findall(r'<script(?! src)(.*?)>(.*?)</script>', h, re.S)
app_js = scripts[-1][1]
with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8') as f:
    f.write(app_js); fname = f.name
r = subprocess.run(['node', '--check', fname], capture_output=True, text=True)
os.unlink(fname)
print("JS:", "OK" if r.returncode==0 else r.stderr[:300])

with open('COCKPIT-ABRIR-AQUI.html', 'w', encoding='utf-8') as f:
    f.write(h)
```

---

## 11. Pitfalls Conhecidos

- **SheetJS mini/core** — NÃO usar. Apenas `xlsx.full.min.js` inclui `cptable`. Versão mini deixa `XLSX.utils` undefined.
- **File System Access API** — requer Chrome/Edge 86+. `queryPermission` retorna `"prompt"` após restart do browser (normal). Usar overlay de reconnect, NÃO auto-request.
- **`<meter>` para barras de progresso no Budget** — `<div>` com `width:%` dentro de `<table><td>` não funciona mesmo com `table-layout:fixed`. Usar `<meter>` nativo com CSS `::-webkit-meter-*`.
- **Nós de projeto no Graph** — NÃO armazenar na sheet `graph_nodes`. São derivados de `projects` via `getEffectiveGraphNodes()`. Posições x/y são salvas em `gNodes` como referência quando o usuário arrasta.
- **Parsing numérico** — sempre usar `parseFloat(String(v||"0").replace(/[^0-9.]/g,""))||0` para valores vindos do Excel. SheetJS pode devolver números como string.
- **Template literals aninhados** — expressões ternárias complexas dentro de `${...}` dentro de template literals dentro de `.map()` dentro de outro template literal podem ter comportamento inesperado. Preferir calcular variáveis antes do template.
- **`applyData`** — sempre chamar `seedAll()` antes de `applyData()`. `applyData` só sobrescreve campos não-vazios.

---

## 12. Backlog / Próximas Features

- [ ] **Dashboard executivo** — view resumida de KPIs de todas as abas numa única tela
- [ ] **Export PDF** — snapshot do cockpit para apresentações
- [ ] **FUP + IA** — usar os registros informativos como contexto para análise automática
- [ ] **Budget — drill-down por projeto** — vincular itens de custo a projetos do Portfolio
- [ ] **Notificações FUP** — alertas visuais para prazos vencidos na nav
- [ ] **Graph — filtro multi-tipo** — selecionar múltiplos tipos de nó simultaneamente
- [ ] **Histórico de versões** — snapshots datados do Excel para tracking de evolução
