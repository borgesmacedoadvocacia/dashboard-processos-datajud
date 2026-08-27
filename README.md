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

1. **Monta a base** — números de processo da planilha-mãe *Clientes e Processos*
   (~2.950 processos) somados aos que já estão na `Base geral`.
2. **Consulta o DataJud** em lote, agrupando os processos por tribunal (o endpoint
   é deduzido do próprio número CNJ). Preenche tribunal, vara, grau, classe, assunto,
   órgão julgador, ajuizamento, sistema, formato, sigilo e **todas as movimentações**.
3. **Varre o DJEN** pelas OABs do escritório e guarda **todas as publicações**,
   sem repetir (dedupe pelo `id` da comunicação, na aba `_DJEN`).
4. **Descobre processos novos** — o que apareceu no DJEN e não estava na base entra
   como linha nova e vai ao DataJud na mesma execução.
5. **Recalcula os indicadores de gestão** (fase, dias parado, situação, últimos eventos).

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

**Tipo de Processo (coluna A)** — o DataJud devolve um registro por grau. A classificação
percorre os registros do mais recente para o mais antigo e adota o primeiro que caia em
um dos quatro tipos; se nenhum cair, é *Processo Principal*. Assim, um processo cuja
1ª instância já virou cumprimento de sentença é classificado como tal, mesmo que exista
um registro antigo de apelação.

Abas de apoio criadas pelo motor: `_Config` (parâmetros), `_Sync` (estado da
sincronização, lido pelo dashboard) e `_DJEN` (publicações brutas, uma por linha).

---

## Instalação do motor (uma vez)

1. Abra a planilha › **Extensões › Apps Script**.
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
| `OABS` | OABs do escritório usadas na varredura do DJEN — `41438/BA, 63805/BA`. Acrescente as demais para ampliar a cobertura. |
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
- **DJEN** limita ~20 requisições por minuto; a varredura por OAB (500 itens por página)
  contorna o limite. Publicações anteriores ao DJEN (2023) não existem na fonte.
- **Sheets** aceita no máximo 50 mil caracteres por célula; processos muito longos têm as
  movimentações/publicações mais antigas omitidas, com aviso no fim da célula.
