# Processos [DataJud] — BM Advocacia

Dashboard estratégico da carteira processual, alimentado automaticamente pela
**API Pública do DataJud (CNJ)** e pelo **DJEN — Diário de Justiça Eletrônico Nacional**.

- **Painel:** https://borgesmacedoadvocacia.github.io/ › menu *Processos [DataJud]*
- **Planilha:** [Processos [DataJud + DJEN]](https://docs.google.com/spreadsheets/d/1MXpGlo3U1bDV34l_jj98ZNPwgguvD-hCFOCaARNQaYs/edit) · aba `Base geral`

---

## Como funciona

```
Clientes e Processos (planilha-mãe)  ─┐
                                      ├─►  Apps Script  ──►  Base geral  ──►  Dashboard
DJEN (OABs do escritório)  ───────────┘      (motor)         (planilha)        (leitura)
DataJud (API pública CNJ)  ───────────┘
```

O motor roda **todo dia às 6h** e **sempre que o botão “Atualizar” é clicado** no
dashboard. Cada execução:

1. **Monta a base** — a lista de números que já está na `Base geral` **manda**
   (hoje, 2.714 processos curados). A planilha-mãe *Clientes e Processos* é usada
   para **enriquecer** partes, cliente e valor da causa; ela só semeia números novos
   quando a aba está vazia (ajustável em `INCLUIR_PLANILHA_MAE`).
2. **Consulta o DataJud** em lote, agrupando os processos por tribunal (o endpoint
   é deduzido do próprio número CNJ). Preenche tribunal, vara, grau, classe, assunto,
   órgão julgador, ajuizamento, sistema, formato, sigilo e **todas as movimentações**.
3. **Varre o DJEN** pelas OABs do escritório e guarda **todas as publicações**,
   sem repetir (dedupe pelo `id` da comunicação, na aba `_DJEN`). Na sincronização
   *completa* — a das 6h e a do botão *Atualizar* — a varredura cobre **todo o
   período** (desde `DJEN_DATA_INICIAL`), não só os últimos dias.
4. **Descobre processos novos** — o que apareceu no DJEN e não estava na base entra
   como linha nova e vai ao DataJud na mesma execução.
5. **Recalcula os indicadores de gestão** (fase, dias parado, situação, últimos eventos).

**A planilha é sempre sobrescrita.** Cada pesquisa regrava as colunas que vêm das
fontes (A, D–O e P–Y) com o resultado da consulta daquele momento — nada de dado
velho sobrevivendo ao lado de dado novo. Se o DataJud responde e o processo não
existe mais lá, as colunas dele são **limpas**; se a consulta em si falhou
(erro de rede ou do CNJ), a linha fica intocada até a próxima rodada. As colunas
preenchidas à mão ou vindas da planilha-mãe (C, Z, AA) nunca são sobrescritas.

Execuções longas se retomam sozinhas: o script grava um cursor e cria um gatilho de
continuação antes de estourar o limite de 6 minutos do Apps Script.

---

## Colunas da aba `Base geral`

| Col | Campo | Origem |
|-----|-------|--------|
| A | Tipo de Processo | classificado pela classe CNJ: **Processo Principal**, **Cumprimento de Sentença**, **Agravo de Instrumento** ou **Recurso de Apelação** |
| B | Número do Processo | base / DJEN |
| C | Partes | planilha-mãe (autor × réu) ou destinatários do DJEN |
| D–M | Tribunal, Vara, Grau, Classe, Assunto, Órgão Julgador, Ajuizamento, Sistema, Formato, Sigilo | DataJud |
| **N** | **Movimentos** | **todas as movimentações do DataJud**, mais recente primeiro, com o grau entre colchetes |
| **O** | **Publicações no DJEN** | **todas as publicações do processo**, com data, tribunal, tipo, órgão e texto integral |
| P–Y | Última Atualização, Qtd. Movimentos, Data/Última Movimentação, Dias sem Movimentação, Qtd./Data/Última Publicação, Fase Processual, Situação | calculadas |
| Z–AB | Valor da Causa, Cliente / Parte Representada, Origem do Cadastro | planilha-mãe |
| AC–AF | Resultado da Sentença, Resultado do Recurso, Fase Processual (cadastro), Situação do Alvará | espelho da planilha-mãe, reescrito a cada sincronização |

**Cliente** é o *nome* de quem o escritório representa — `AUTOR` ou `RÉU` conforme a coluna
PARTE REPRESENTADA da planilha-mãe (que guarda o polo, não o nome).

**Duas fases.** `Fase Processual` (X) é inferida das movimentações do DataJud; `Fase Processual
(cadastro)` (AE) é a classificação do escritório. A divergência entre as duas é justamente o
sinal de que o cadastro ficou para trás — o consultor I.A. sabe comparar as duas.

**Tipo de Processo (coluna A)** — o DataJud devolve um registro por grau. A classificação
percorre os registros do mais recente para o mais antigo e adota o primeiro que caia em
um dos quatro tipos; se nenhum cair, é *Processo Principal*. Assim, um processo cuja
1ª instância já virou cumprimento de sentença é classificado como tal, mesmo que exista
um registro antigo de apelação.

Abas de apoio criadas pelo motor: `_Config` (parâmetros), `_Sync` (estado da
sincronização, lido pelo dashboard) e `_DJEN` (publicações brutas, uma por linha).

---

## Instalação do motor (uma vez)

1. Abra a planilha › **Extensões › Apps Script** (script vinculado — é o que faz
   aparecer o menu *BM · DataJud/DJEN* dentro da planilha). Um projeto **independente**
   criado direto no script.google.com também funciona: nesse caso `SpreadsheetApp.getActive()`
   volta `null` e o motor abre a planilha por `PLANILHA_ID` — só não haverá menu na planilha,
   as funções são executadas pelo seletor de função do editor.
2. Crie dois arquivos e cole o conteúdo de `apps-script/Codigo.gs` e `apps-script/Fontes.gs`.
3. Rode a função **`configurarTudo`** e autorize os acessos pedidos.
   Isso cria as abas de apoio, ajusta o cabeçalho e agenda o gatilho das **6h**.
4. **Implantar › Nova implantação › App da Web**
   - *Executar como:* **Eu**
   - *Quem pode acessar:* **Qualquer pessoa**
   - Copie a URL que termina em `/exec`.
5. No dashboard, clique em **⚙** e cole essa URL e o token
   (campo `TOKEN_WEBAPP` da aba `_Config` — troque o valor padrão).
6. Primeira carga: menu **BM · DataJud/DJEN › Sincronizar agora (completa)**.
   A varredura histórica do DJEN (desde 2023) leva alguns minutos e se retoma sozinha.

### Parâmetros (aba `_Config`)

| Chave | Para que serve |
|-------|----------------|
| `PLANILHA_ID` | ID da planilha de destino. Só é usado quando o script é um projeto independente. |
| `INCLUIR_PLANILHA_MAE` | `SO_SE_VAZIA` (padrão — preserva a lista curada), `SEMPRE` (junta os ~2.950 processos da planilha-mãe) ou `NUNCA`. Em qualquer opção a planilha-mãe enriquece partes, cliente e valor da causa. |
| `OABS` | OABs varridas no DJEN — `41438/BA, 271081/RJ, 536843/SP, 63805/BA`. Aceita `41438/BA`, `BA/41438` ou `OAB/BA 41438`; duplicatas são ignoradas. Depois de mexer no padrão do código, rode **Aplicar OABs padrao no _Config** (o `_Config` tem precedência sobre o código). |
| `DJEN_DATA_INICIAL` | Início da varredura histórica (padrão `2023-01-01`). |
| `DJEN_JANELA_DIAS` | Janela das varreduras seguintes (padrão 45 dias). |
| `DESCOBRIR_NO_DJEN` | `SIM` inclui na base processos que aparecem no DJEN e não estão na planilha-mãe. |
| `MAX_MOVIMENTOS` / `MAX_PUBLICACOES` | Teto por processo (limite de 50 mil caracteres por célula do Sheets). |
| `TOKEN_WEBAPP` | Senha do botão *Atualizar*. |
| `DIAS_ESTAGNADO` | Dias sem movimento para marcar o processo como *Estagnado* (padrão 90). |

---

## Consultor I.A.

Campo de perguntas no próprio dashboard, usando **Claude Sonnet 5** (`claude-sonnet-5`)
com *tool use* sobre a planilha: o navegador executa as consultas e devolve os números
ao modelo, que só responde com o que foi apurado — nada é estimado.

Ferramentas disponíveis ao consultor: `consultar_base` (filtros, agrupamentos e amostras),
`ler_processo` (movimentações e publicações integrais de um processo) e `panorama`
(visão executiva da carteira inteira).

A chave da API fica **apenas no navegador** (`localStorage`), nunca no repositório.

---

## Limites conhecidos

- **DataJud não traz nomes das partes** (dado pessoal). Partes e cliente vêm da planilha-mãe.
- **Cobertura do DataJud** gira em torno de 85–90% dos processos: números em segredo de
  justiça, tribunais que não alimentam a base e números digitados errado ficam como
  *“Sem dados nas fontes”* — o dashboard tem um alerta próprio para eles.
- **OAB/RJ 271081** existe, mas hoje não tem nenhuma publicação no DJEN (conferido em 27/08/2026,
  período completo desde 2023). Fica configurada assim mesmo: se passar a receber, entra sozinha.
- **DJEN** limita ~20 requisições por minuto; a varredura por OAB (500 itens por página)
  contorna o limite. Publicações anteriores ao DJEN (2023) não existem na fonte.
- **Sheets** aceita no máximo 50 mil caracteres por célula; processos muito longos têm as
  movimentações/publicações mais antigas omitidas, com aviso no fim da célula.
