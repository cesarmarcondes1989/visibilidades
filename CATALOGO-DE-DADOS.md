# Catálogo de Dados — Cockpit Estratégico de Dados

> Documento de referência do modelo de dados do app `COCKPIT-ABRIR-AQUI27.html`.
> Gerado para uso com Copilot / assistentes de IA, para que entendam exatamente
> como os dados são estruturados, persistidos e relacionados.

## 1. Visão geral do armazenamento

- O app é um **HTML único** (sem backend, sem banco de dados). Roda 100% no navegador (Chrome/Edge 86+).
- Os dados são persistidos em um arquivo **Excel local** chamado `estrategia.xlsx`, lido/escrito via **File System Access API** do navegador. A leitura/escrita do `.xlsx` é feita com a biblioteca **SheetJS**, embutida no próprio HTML.
- **Cada aba (sheet) do Excel = uma tabela.** Cada linha = um registro. A primeira linha de cada aba é o cabeçalho (nomes das colunas).
- Existem dois "formatos" de aba:
  1. **Tabelas linha-a-coluna** (a maioria): uma linha por registro, colunas fixas. Construídas por `_list2ws(items, cols)` e lidas por `_list(wb, nome)`.
  2. **Abas chave-valor** (`_vision` e `_data_vision`): apenas duas colunas, `chave` e `valor`, representando um objeto único (não uma lista de registros). Construídas por `_kv2ws(obj)` e lidas por `_kv(wb, nome)`.
- Ao salvar, valores que são array/objeto JS são serializados como **string JSON** dentro da célula (ex.: a coluna `valor` da aba `_vision` pode conter `[{"text":"..."}]`). Ao ler, strings que começam com `[` ou `{` são automaticamente parseadas de volta para JSON; strings `"TRUE"/"FALSE"` (qualquer caixa) são convertidas para booleano.
- **Merge-on-load**: ao carregar o Excel, cada lista só substitui os dados em memória **se vier não-vazia** (`if (o.bets?.length) bets = o.bets`). Ou seja, uma aba vazia/ausente no Excel não apaga os dados já carregados em memória — ela simplesmente é ignorada nesse load.
- Botão "Reset" recarrega todos os dados-padrão (`seedAll()`), descartando os dados atuais em memória (não afeta o arquivo Excel até salvar).

## 2. Mapa de tabelas (abas do Excel)

| Aba (sheet) | Tipo | Entidade representada | Tem ID próprio? |
|---|---|---|---|
| `_vision` | chave-valor | Visão estratégica (singleton) | — |
| `themes` | lista | Temas estratégicos (sub-coleção de `_vision.themes`) | não (índice no array) |
| `_data_vision` | chave-valor | Visão de dados (singleton) | — |
| `principles` | lista | Princípios de dados (sub-coleção de `_data_vision.principles`) | não (índice no array) |
| `bets` | lista | Apostas estratégicas (Bets) | sim — `b<timestamp>` |
| `data_products` | lista | Produtos de dados | sim — `dp<timestamp>` |
| `dp_subs` | lista | Subprojetos de um produto de dados | sim — `dps<timestamp>` |
| `maturity` | lista | Dimensões de maturidade de dados | sim — `catalog`/`quality`/... (string fixa do seed) ou sem novo-cadastro pela UI |
| `governance` | lista | Capacidades de governança | não (índice no array) |
| `projects` | lista | Projetos do portfólio | sim — `proj<timestamp>` |
| `melhorias` | lista | Iniciativas de melhoria (Salesforce / Operacional) | sim — `m<timestamp>` |
| `graph_nodes` | lista | Nós do grafo **(apenas tipos pilar/produto/dominio)** | sim — `gn<timestamp>` ou prefixos do seed (`np-`, `pr-`, `d-`) |
| `graph_edges` | lista | Conexões (arestas) do grafo | não (chave composta `from`+`to`) |
| `pillars` | lista | Pilares estratégicos (catálogo mestre) | sim — slug textual, ex. `mro` |
| `domains` | lista | Domínios de dados (catálogo mestre) | sim — slug textual, ex. `health` |
| `budget` | lista | Itens de orçamento/custo | sim — `bg<timestamp>` |
| `fup` | lista | Itens de acompanhamento (Follow-up) | sim — `fp<timestamp>` |
| `people` | lista | Pessoas do time | sim — `pe<timestamp>` |
| `people_skills` | lista | Skills associadas a uma pessoa | sim — `psk<timestamp>` |
| `people_achievements` | lista | Achievements associados a uma pessoa | sim — `pach<timestamp>` |
| `prateleira` | lista | Itens de "prateleira" (catálogo de recursos/custos reutilizáveis) | sim — `pt<timestamp>` |

> ⚠️ **Nós derivados/virtuais**: nós do tipo `projeto` e `pessoa` **NÃO** existem na aba `graph_nodes` — eles são calculados em tempo de execução a partir das tabelas `projects` e `people` pela função `getEffectiveGraphNodes()`. O ID do nó derivado de um projeto é `"pj-" + projeto.id`; o de uma pessoa é `"pe-" + pessoa.id`. Ao reconstruir/ler o Excel, a aba `graph_nodes` é explicitamente filtrada para excluir esses dois tipos. Qualquer leitor do dataset (incluindo Copilot) deve tratar a lista completa de nós do grafo como **`graph_nodes` (pilar/produto/dominio) + nós derivados de `projects` + nós derivados de `people`**, nunca apenas `graph_nodes`.

## 3. Detalhamento por tabela

### 3.1 `_vision` (chave-valor — singleton "Visão Estratégica")

Não é uma lista de registros; é um único objeto com estas chaves:

| Chave | Tipo | Descrição |
|---|---|---|
| `intent` | texto | Intenção estratégica (frase-síntese da visão) |
| `northStar` | texto | Nome da métrica North Star |
| `nsCurrentPct` | número | Valor atual da North Star (%) |
| `nsTargetPct` | número | Valor-alvo da North Star (%) |
| `nsCurrentLabel` | texto | Rótulo de exibição do valor atual (ex.: "8% today") |
| `nsTargetLabel` | texto | Rótulo de exibição do valor-alvo (ex.: "30% target (2028)") |
| `diagPositive` | array de strings (JSON) | Lista de "forças" do diagnóstico estratégico |
| `diagNegative` | array de strings (JSON) | Lista de "gaps" do diagnóstico estratégico |
| `themes` | — | **Não fica na aba `_vision`**: é extraído e persistido separadamente na aba `themes` (ver 3.2). Ao ler, é reanexado a este objeto como `vision.themes`. |

### 3.2 `themes` (sub-coleção de `_vision`)

Colunas: `name`, `desc`, `color`.

Representa os "temas estratégicos" da visão. Não é uma tabela independente — é exibida e editada dentro da seção de Visão, sempre como sub-lista de `vision.themes`. Cada item: `{name, desc, color}` (color = hex, escolhido de uma paleta fixa: `#0DAF7B #4F70F5 #7C3AED #D97706 #DB2777 #0891B2 #059669`). Sem ID próprio — editado/excluído por índice no array.

### 3.3 `_data_vision` (chave-valor — singleton "Visão de Dados")

| Chave | Tipo | Descrição |
|---|---|---|
| `statement` | texto | Declaração da estratégia de dados |
| `principles` | — | **Não fica na aba `_data_vision`**: persistido separadamente na aba `principles` (ver 3.4). |

### 3.4 `principles` (sub-coleção de `_data_vision`)

Colunas: `icon`, `name`, `desc`.

Princípios de dados (ex.: "Data as product", "API-first"). `icon` é um emoji. Sem ID próprio — editado/excluído por índice.

### 3.5 `bets` — Apostas estratégicas

Colunas (`SH.bets`): `id, hz, pillar, name, intent, porque, enablerDados, enablerTech, metrica, meta, atual, status`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | `"b"+Date.now()` |
| `hz` | enum | horizonte: `h1`\|`h2`\|`h3` → FK lógica para `HZ` (não é tabela persistida, é constante do app) |
| `pillar` | string | FK lógica → `pillars.id` |
| `name` | texto | Nome da aposta |
| `intent` | texto longo | Intenção estratégica |
| `porque` | texto longo | Racional / "why now" |
| `enablerDados` | texto longo | Dados necessários como enabler |
| `enablerTech` | texto longo | Tecnologia necessária como enabler |
| `metrica` | texto | Nome da métrica de sucesso |
| `meta` | texto | Valor-alvo da métrica (texto livre, pode ter "%", "R$" etc.) |
| `atual` | texto | Valor atual da métrica (texto livre) |
| `status` | enum | um dos valores de `STLIST` (ver §5) |

### 3.6 `data_products` — Produtos de dados

Colunas (`SH.data_products`): `id, domain, name, owner, q, cons, status, horizon, sla, notas`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | `"dp"+Date.now()` |
| `domain` | string | FK lógica → `domains.id` |
| `name` | texto | Nome do produto de dados |
| `owner` | texto | Time/pessoa responsável |
| `q` | número 0–100 | Score de "Quality" (qualidade) |
| `cons` | número | Quantidade de consumidores/receita associada |
| `status` | enum | `Production`\|`Beta`\|`Planned`\|`Discontinued` |
| `horizon` | enum | `h1`\|`h2`\|`h3` |
| `sla` | booleano | Se possui SLA/contrato de dados formal |
| `notas` | texto longo | Notas livres |

Cada produto de dados pode ter **subprojetos** (tabela `dp_subs`) e é excluído em cascata: ao deletar um `data_products`, todos os `dp_subs` com `dpId` igual são removidos também.

### 3.7 `dp_subs` — Subprojetos de um produto de dados

Colunas (`SH.dp_subs`): `id, dpId, name, descr, status, prog, start, end`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | `"dps"+Date.now()` |
| `dpId` | string | FK → `data_products.id` |
| `name` | texto | Nome do subprojeto |
| `descr` | texto longo | Descrição |
| `status` | enum | um dos valores de `STLIST` (ver §5) |
| `prog` | número 0–100 | Progresso (%) |
| `start` | data (ISO `YYYY-MM-DD`) | Início |
| `end` | data (ISO `YYYY-MM-DD`) | Fim |

### 3.8 `maturity` — Dimensões de maturidade de dados

Colunas (`SH.maturity`): `id, name, atual, alvo, owner`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | slug fixo do seed (ex.: `catalog`, `quality`, `lineage`, `privacy`, `owners`, `sla`, `obs`) |
| `name` | texto | Nome da dimensão |
| `atual` | número 0–5 | Maturidade atual |
| `alvo` | número 0–5 | Meta de maturidade |
| `owner` | texto | Responsável |

### 3.9 `governance` — Capacidades de governança

Colunas (`SH.governance`): `name, atual, alvo, owner, note`.

Sem ID próprio — editado por índice no array. `atual`/`alvo` são escalas de maturidade (range usado no seed é 0–5, mas o slider de edição (`sldr`) não impõe limite fixo de UI nesta tela — confirme com `m.atual`/`m.alvo`). `note` é texto livre ("next steps").

### 3.10 `projects` — Projetos do portfólio

Colunas (`SH.projects`): `id, hz, pillar, bet, name, descr, inv, rec, prog, status, start, end`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | `"proj"+Date.now()` |
| `hz` | enum | `h1`\|`h2`\|`h3` |
| `pillar` | string | FK lógica → `pillars.id` |
| `bet` | string | FK lógica → `bets.id` (aposta vinculada) |
| `name` | texto | Nome do projeto |
| `descr` | texto longo | Descrição |
| `inv` | número (R$) | Investimento |
| `rec` | número (R$) | Retorno esperado |
| `prog` | número 0–100 | Progresso (%) |
| `status` | enum | um dos valores de `STLIST` (ver §5) |
| `start` | data (ISO) | Início |
| `end` | data (ISO) | Fim |

Cada projeto também gera, em tempo de execução, um **nó de grafo derivado** (`pj-`+id, tipo `projeto`) — ver §2 nota sobre nós virtuais.

### 3.11 `melhorias` — Iniciativas de melhoria

Colunas (`SH.melhorias`): `id, area, name, status, inv, rec, prog, notas`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | `"m"+Date.now()` |
| `area` | enum | `sf` (Salesforce) \| `mel` (Operational Improvement) |
| `name` | texto | Nome da iniciativa |
| `status` | enum | ⚠️ **lista própria, diferente de `STLIST`**: `In Progress`\|`Planned`\|`Done`\|`At Risk`\|`Paused` (ver nota em §5) |
| `inv` | número (R$) | Investimento |
| `rec` | número (R$) | Retorno esperado |
| `prog` | número 0–100 | Progresso (%) |
| `notas` | texto longo | Notas livres |

### 3.12 `graph_nodes` — Nós do grafo (somente pilar/produto/dominio)

Colunas (`SH.graph_nodes`): `id, type, label, group`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | prefixo por tipo: `np-` (pilar), `pr-` (produto), `d-` (dominio). Novos nós criados manualmente: `"gn"+Date.now()` |
| `type` | enum | uma das chaves de `NODE_TYPES` — nesta aba, apenas `pilar`\|`produto`\|`dominio` (tipos `projeto`/`pessoa` são excluídos, ver §2) |
| `label` | texto | Rótulo exibido no grafo |
| `group` | texto | Agrupamento visual (opcional; default = `type`) |

Nós também podem ter posição salva (`x`, `y`) em memória, usada pelo layout do grafo, mas isso é tratado como estado de layout, não uma coluna persistida explicitamente listada em `SH.graph_nodes` (a posição não é serializada nesta aba).

### 3.13 `graph_edges` — Conexões do grafo

Colunas (`SH.graph_edges`): `from, to, intensity, label`.

| Coluna | Tipo | Observações |
|---|---|---|
| `from` | string | ID do nó de origem (qualquer prefixo: `np-`,`pr-`,`pj-`,`d-`,`pe-`) |
| `to` | string | ID do nó de destino |
| `intensity` | número 1–5 | Força da conexão (1=fraca, 5=forte) |
| `label` | texto opcional | Rótulo da relação (ex.: `strategy`, `data dependency`, `delivery`, `technical dependency`) |

Sem ID próprio — uma aresta é identificada pelo par `(from, to)`.

### 3.14 `pillars` — Catálogo mestre de pilares estratégicos

Colunas: `id, name, color` (lista fixa, não está em `SH` pois usa um array de colunas inline em `buildWB()`).

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | slug único, ex.: `mro`, `saude`, `ops`, `dados`, `mat`, `ia`, `novos` |
| `name` | texto | Nome de exibição |
| `color` | hex | Cor associada |

Seed inicial = constante `PILLARS` (7 pilares). Editável via tela Admin. Exclusão verifica uso em `bets`/`projects` antes de remover (apenas avisa, não bloqueia).

### 3.15 `domains` — Catálogo mestre de domínios de dados

Colunas: `id, name, color` (mesmo padrão de `pillars`).

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | slug único, ex.: `health`, `mro`, `ops`, `parts`, `cust`, `eng`, `fin` |
| `name` | texto | Nome de exibição |
| `color` | hex | Cor associada |

Seed inicial = constante `DOMAINS` (7 domínios). Exclusão verifica uso em `data_products` antes de remover (apenas avisa).

### 3.16 `budget` — Itens de orçamento

Colunas (`SH.budget`): `id, name, categoria, anual, ytd, notas`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | `"bg"+Date.now()` |
| `name` | texto | Nome do item de custo |
| `categoria` | texto livre | ex.: "Pessoal", "Licenças", "Cloud" (sem enum fixo — agrupamento dinâmico na UI) |
| `anual` | número (R$) | Valor anual planejado |
| `ytd` | número (R$) | Valor realizado no ano (year-to-date) |
| `notas` | texto longo | Notas livres |

### 3.17 `fup` — Itens de acompanhamento (Follow-up)

Colunas (`SH.fup`): `id, tipo, titulo, responsavel, prazo, prioridade, status, acompanhar, notas`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | `"fp"+Date.now()` |
| `tipo` | enum | `acomp` (atividade de acompanhamento, com todos os campos) \| `info` (registro informativo — usa só `titulo`, `prazo`/data e `notas`) |
| `titulo` | texto | Título (obrigatório) |
| `responsavel` | texto | Apenas para `tipo=acomp` |
| `prazo` | data (ISO) | Para `acomp` é "prazo"; para `info` é apenas "data" |
| `prioridade` | enum | Apenas para `tipo=acomp`: `Alta`\|`Média`\|`Baixa` |
| `status` | enum | Apenas para `tipo=acomp`: `Pendente`\|`Em andamento`\|`Concluído`\|`Bloqueado`\|`Cancelado` |
| `acompanhar` | booleano (string `"true"/"false"`) | Apenas para `tipo=acomp` — flag "marcar para acompanhar" |
| `notas` | texto longo | Notas livres |

### 3.18 `people` — Pessoas do time

Colunas (`SH.people`): `id, name, cargo`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | `"pe"+Date.now()` |
| `name` | texto | Nome (obrigatório) |
| `cargo` | texto | Cargo/título |

Exclusão em cascata: ao remover uma pessoa, remove também seus `people_skills`, `people_achievements`, e quaisquer `graph_edges`/nó de grafo derivado (`pe-`+id) associados.

Cada pessoa também gera, em tempo de execução, um **nó de grafo derivado** (`pe-`+id, tipo `pessoa`) — ver §2.

### 3.19 `people_skills` — Skills de uma pessoa

Colunas (`SH.people_skills`): `id, personId, skill`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | `"psk"+Date.now()` |
| `personId` | string | FK → `people.id` |
| `skill` | texto | Texto livre da skill (ex.: "SQL", "Liderança", "Python") |

### 3.20 `people_achievements` — Achievements de uma pessoa

Colunas (`SH.people_achievements`): `id, personId, title, descr, date`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | `"pach"+Date.now()` |
| `personId` | string | FK → `people.id` |
| `title` | texto | Título do achievement (obrigatório) |
| `descr` | texto longo | Descrição |
| `date` | data (ISO) | Data do achievement |

### 3.21 `prateleira` — Catálogo de itens/recursos reutilizáveis

Colunas (`SH.prateleira`): `id, name, categoria, valor, moeda, unidade, notas`.

| Coluna | Tipo | Observações |
|---|---|---|
| `id` | string | `"pt"+Date.now()` |
| `name` | texto | Nome do item (ex.: "Recurso Engenharia") |
| `categoria` | texto livre | ex.: "Pessoal", "Serviço", "Licença" |
| `valor` | número | Valor monetário |
| `moeda` | enum | `BRL`\|`USD`\|`EUR` (símbolos: R$, US$, €) |
| `unidade` | enum | `hora`\|`dia`\|`mes`\|`ano`\|`projeto`\|`unico` (rótulos: "Por hora", "Por dia", "Por mês", "Por ano", "Por projeto", "Valor único") |
| `notas` | texto longo | Notas livres |

## 4. Relacionamentos (chaves lógicas)

Não há chaves estrangeiras de banco de dados (Excel não impõe integridade referencial) — todas as relações abaixo são por convenção de valor de string, validadas/usadas apenas em memória pelo app:

- `bets.pillar` → `pillars.id`
- `projects.pillar` → `pillars.id`
- `projects.bet` → `bets.id`
- `data_products.domain` → `domains.id`
- `dp_subs.dpId` → `data_products.id`
- `people_skills.personId` → `people.id`
- `people_achievements.personId` → `people.id`
- `graph_edges.from` / `graph_edges.to` → ID de nó do grafo (qualquer nó efetivo: `graph_nodes` + nós derivados de `projects`/`people`)
- `bets.hz`, `projects.hz`, `data_products.horizon` → `id` de `HZ` (constante, não persistida como tabela)

### Convenção de prefixos de ID de nó de grafo

| Prefixo | Tipo de nó | Origem |
|---|---|---|
| `np-` | `pilar` | seed manual / `graph_nodes` |
| `pr-` | `produto` | seed manual / `graph_nodes` |
| `d-` | `dominio` | seed manual / `graph_nodes` |
| `pj-` | `projeto` | **derivado** de `projects.id` (não persistido em `graph_nodes`) |
| `pe-` | `pessoa` | **derivado** de `people.id` (não persistido em `graph_nodes`) |

## 5. Enumerações / valores permitidos

| Enum | Valores | Usado em |
|---|---|---|
| `STLIST` (status genérico) | `Active`, `Bet`, `Exploration`, `Done`, `Paused`, `At Risk` | `bets.status`, `dp_subs.status`, `projects.status` |
| Status de `melhorias` ⚠️ **lista diferente** | `In Progress`, `Planned`, `Done`, `At Risk`, `Paused` | `melhorias.status` — **não confundir com `STLIST`**, embora pareçam similares, os valores e o significado de cada item diferem (não tem `Exploration`, e usa `In Progress` em vez de `Active`) |
| Status de `data_products` | `Production`, `Beta`, `Planned`, `Discontinued` | `data_products.status` |
| Status de `fup` (tipo=acomp) | `Pendente`, `Em andamento`, `Concluído`, `Bloqueado`, `Cancelado` | `fup.status` |
| Prioridade de `fup` | `Alta`, `Média`, `Baixa` | `fup.prioridade` |
| Tipo de `fup` | `acomp` (acompanhamento), `info` (informativo) | `fup.tipo` |
| Área de `melhorias` | `sf` (Salesforce), `mel` (Operational Improvement) | `melhorias.area` |
| Moeda de `prateleira` | `BRL`, `USD`, `EUR` | `prateleira.moeda` |
| Unidade de `prateleira` | `hora`, `dia`, `mes`, `ano`, `projeto`, `unico` | `prateleira.unidade` |
| `NODE_TYPES` (tipos de entidade do grafo) | `pilar`, `produto`, `projeto`, `dominio`, `pessoa` | `graph_nodes.type` + nós derivados |
| `HZ` (horizontes estratégicos) | `h1` (Now, 0–12 meses), `h2` (Scale, 1–3 anos), `h3` (Differentiation, 3–5+ anos) | `bets.hz`, `projects.hz`, `data_products.horizon` |
| `PILLARS` (seed inicial — 7 pilares) | `mro` (Connected MRO), `saude` (Health & Predictive), `ops` (Flight Operations), `dados` (Data Platform), `mat` (Aftermarket), `ia` (AI & GenAI), `novos` (New Markets) | tabela `pillars` (editável) |
| `DOMAINS` (seed inicial — 7 domínios) | `health` (Aircraft Health), `mro` (MRO & Maintenance), `ops` (Flight Operations), `parts` (Parts & Supply), `cust` (Customer & Commercial), `eng` (Engineering & Quality), `fin` (Financial) | tabela `domains` (editável) |

## 6. Convenções de ID por entidade

| Entidade | Padrão de ID | Exemplo |
|---|---|---|
| Bets | `"b"+Date.now()` | `b1700000000000` |
| Data Products | `"dp"+Date.now()` | |
| dp_subs | `"dps"+Date.now()` | |
| Projects | `"proj"+Date.now()` | |
| Melhorias | `"m"+Date.now()` | |
| Budget | `"bg"+Date.now()` | |
| Prateleira | `"pt"+Date.now()` | |
| FUP | `"fp"+Date.now()` | |
| People | `"pe"+Date.now()` | |
| people_skills | `"psk"+Date.now()` | |
| people_achievements | `"pach"+Date.now()` | |
| Graph nodes (manuais) | `"gn"+Date.now()` | |
| Pillars | slug textual definido pelo usuário | `mro` |
| Domains | slug textual definido pelo usuário | `health` |
| Governance | sem ID — referenciado por índice no array | — |
| Maturity (novos itens) | sem criação pela UI no fluxo atual — itens do seed têm slug fixo | `catalog` |
| Themes / Principles | sem ID — referenciado por índice no array | — |

> Nota: como os IDs usam `Date.now()` (milissegundos), criar dois registros da mesma entidade no mesmo milissegundo geraria colisão — cenário extremamente improvável em uso manual via UI.

## 7. Notas para uso com Copilot / IA

- Ao perguntar sobre "o que está na planilha X", mapeie para a tabela correspondente na seção 3 acima.
- Se for perguntado sobre nós/conexões do grafo estratégico, lembre-se de que a lista completa de nós é `graph_nodes` + nós derivados de `projects` (prefixo `pj-`) + nós derivados de `people` (prefixo `pe-`) — nunca apenas a aba `graph_nodes` isolada.
- `themes` e `principles` não são entidades de topo independentes — são sub-listas de `_vision` e `_data_vision`, respectivamente, mesmo estando em abas separadas no Excel.
- Cuidado ao comparar status entre `melhorias` e as demais entidades (`bets`/`projects`/`dp_subs`) — usam listas de enum diferentes (ver §5).
- Datas são sempre strings no formato ISO `YYYY-MM-DD` (inputs `type="date"` do HTML).
- Valores monetários (`inv`, `rec`, `anual`, `ytd`, `valor`) são números (R$ por padrão, exceto `prateleira.valor` que tem moeda própria via `moeda`).
