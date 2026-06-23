# Critérios da simulação final

## Ficheiro base e saída

- **Template:** `ficheiro template 2.xlsx`
- **Saída:** `SXX_analise_precos_SUIVI.xlsx`, em que `XX` é o número da semana da simulação
  mais recente.

## Deteção automática de ficheiros

| Origem | Regra |
|---|---|
| Simulação PVP | Ficheiro `*Simulação PVP*` com o **S mais alto** no nome |
| Análise preços (histórico) | `S(XX-1) - analise preços.xlsx` |
| Relatório comparável | Ficheiro `*compar*` com a **data YYYYMMDD** mais recente no nome |
| Total meas | Ficheiro `TOTAL*meas*` com a **data DD-MM-YYYY** mais recente no nome |

## Fonte principal (TD Psyco)

- A informação base segue a tabela dinâmica `TD Psyco` da simulação semanal.
- A `TD Psyco` pode estar filtrada no Excel. Para gerar o ficheiro final, devem ser
  considerados todos os dados, como se os filtros da pivot fossem removidos.
- Como a pivot pode estar guardada com cache filtrada, o gerador reconstrói a
  `TD Psyco` sem filtros a partir da folha `Dados`.
- A reconstrução exclui `MERCADONA`, porque o ficheiro final só usa as colunas
  Shopping `CONTINENTE`, `LIDL` e `PINGO-DOCE`.
- Excluir artigos cuja descrição comece por `SUB` ou `XXX`. Artigos `PAL` mantêm-se.

## Colunas copiadas da Simulação

| Origem | Destino no template |
|---|---|
| Nomenclatura Amont | NOMENCLATURA AMON |
| Nomenclatura Aval | NOMENCLATURA AVAL |
| ITM8 | ITM8 |
| EAN | EAN |
| Descrição | DESCRIÇÃO ARTIGO |
| Marca | MARCA |
| Tipo Produto | TIPO |
| Psyco | PSYCHO |
| Argus | ARGUS |
| Promo Permanente | PROMO PERM (LEADER) |
| EDLP | EDLP |
| Estatuto (folha Dados) | ST |
| PVP Cadencier Futuro | PVP CADENCIER FUTURO |
| CONTINENTE | CONTINENTE |
| LIDL | LIDL |
| PINGO-DOCE | PINGO-DOCE |

## Fórmulas derivadas (A/B → C–G)

| Coluna | Fórmula |
|---|---|
| GRP | `=SEG.TEXTO(A;2;2)` |
| GRUPO | `=PROCX(C;Folha2!A:A;Folha2!B:B)` |
| FIL | `=SEG.TEXTO(A;5;3)` |
| FAM | `=SEG.TEXTO(B;5;2)` |
| SFA | `=SEG.TEXTO(B;8;2)` |

## Promo SoySuper (Relatório comparável)

PROCX por EAN (coluna K do template ↔ coluna A do comparável):

| Template | Origem no comparável |
|---|---|
| CNT (V) | col. R — Preço com promoção Continente |
| LIDL (W) | col. Z — Preço com promoção Lidl |
| PD (X) | col. V — Preço com promoção Pingo Doce |

Se não houver valor, a célula fica vazia.

## Fórmulas calculadas (linhas com EAN)

| Coluna | Fórmula |
|---|---|
| CONDIÇÃO PVP (Y) | `=SE(OU(N="O";O="O");SE(CONTAR(V:X)>0;MÍNIMO(V:X);MÍNIMO(S:U));SE(M="O";MÍNIMO(S:U);SE(L="E";SE.ERRO(MODA(S:U);MÍNIMO(S:U));MÉDIA(S:U))))` |
| DESVIO % (Z) | `=R/Y-1` |
| DESVIO € (AA) | `=R-Y` |
| CHECK (AB) | `=SE(R=Y;"VERDADEIRO";"FALSO")` |
| PVP FOLHETO BLQ (AE) | `=SE(AC="";"";SE(OU(AC>=(HOJE()+15);AC<(HOJE()+7));"não";"sim"))` — formato **Geral** |

- Colunas **S:Y** com formato **Moeda €**.

## Total Meas

O ficheiro Total Meas (indicativo de folheto) é primeiro processado em duas folhas auxiliares:

1. **TD Meas** — tabela dinâmica em formato tabular com `GRUPO_INTERNO`, `UVC`,
   `EAN`, `DESCRIÇÃO`, `IN_MEA`, `PVP`, ordenada por Grupo / UVC / EAN /
   Descrição / data mais recente.
2. **Meas Processado** — uma linha por `EAN` com a `IN_MEA` mais recente.

O ficheiro processado é guardado como `TOTAL - meas a DD-MM-YYYY - processado.xlsx`.

A ligação ao template é feita por **EAN**. Nem todos os artigos da TD Psyco têm
folheto, pelo que é normal existirem linhas sem campanha/PVP preenchidos.

| Origem | Destino |
|---|---|
| IN_MEA | CAMPANHA PROMO FUTURA (AC) |
| PVP | PVP CAMPANHA (AD) |

## Análise preços da semana anterior (Simulação − 1)

| Origem (SXX-1) | Destino (novo ficheiro) |
|---|---|
| Comentários (face ao suivi) | HISTORICO (face ao suivi) — AI |
| Comentários comercial | HISTÓRICO (comentários comercial) — AL |

Se a coluna de comentários da semana anterior estiver vazia, usa-se a coluna
`HISTORICO` / `HISTORICO (face ao suivi)` do ficheiro anterior para manter a cadeia.

As colunas **COMENTÁRIOS (face ao suivi)** (AH) e **FEEDBACK COMERCIAL** (AJ) ficam sempre
vazias no ficheiro gerado.

## Listas de validação (Folha2)

À direita do mapa Grp/GRUPO existente, com colunas de intervalo:

- **PROPOSTA** (AG): Subir / Descer / OK
- **FEEDBACK COMERCIAL** (AJ): OK Envio ao ficheiro / NOK / Sem Con / Nego / Acompanhar

## Campos em branco

- **NOVO PVP** (AF)
- **COMENTÁRIOS (face ao suivi)** (AH)
- **FEEDBACK COMERCIAL** (AJ)
- **COMENTÁRIOS COMERCIAL** (AK)

## Etiquetas da linha 2

- `S2`: `Shopping DD/MM/AAAA`
- `V2`: `PROMO DD/MM/AAAA`

Data obtida da simulação (folha `Dados`).

## Ordenação

1. Grp — coluna C
2. FIL — coluna E
3. FAM — coluna F
4. SFA — coluna G
5. MARCA — coluna J
6. PVP CADENCIER FUTURO — coluna R
