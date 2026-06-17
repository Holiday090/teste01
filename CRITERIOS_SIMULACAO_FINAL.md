# Critérios da simulação final

## Fonte principal

- A informação base do ficheiro final deve seguir a tabela dinâmica `TD Psyco` da
  `S24 - Simulação PVP S24-2026.xlsm`.
- A `TD Psyco` pode estar filtrada no Excel. Para gerar o ficheiro final, devem
  ser considerados todos os dados, como se os filtros da pivot fossem removidos.
- Como a pivot pode estar guardada com cache filtrada, o gerador reconstrói a
  `TD Psyco` sem filtros a partir da folha `Dados`.
- A coluna `R` do ficheiro final deve ser igual à coluna `L` da `TD Psyco`
  (`PVP Cadencier Actual`).
- As colunas Shopping `S:T:U` do ficheiro final devem ser iguais às colunas
  `N:O:P` da `TD Psyco`, respetivamente `CONTINENTE`, `LIDL` e `PINGO-DOCE`.
- A reconstrução exclui `MERCADONA`, porque a `TD Psyco`/ficheiro final só usa
  as colunas Shopping `CONTINENTE`, `LIDL` e `PINGO-DOCE`.
- Os artigos cuja descrição comece por `SUB.` ou `PAL` devem ser excluídos do
  ficheiro final.

## Estatuto

- A coluna `ST` do ficheiro final vem sempre da folha `Dados`.

## Ficheiros auxiliares

- Quando não forem indicados explicitamente no comando, o `Relatorio comparavel`
  e o `TOTAL MEAS` devem ser escolhidos automaticamente pelo ficheiro Excel mais
  recente na pasta de trabalho que corresponda a esses nomes.

## Semana anterior e comentários

- A indicação `S23`, `S24`, etc. no nome do ficheiro representa o número da
  semana.
- Cada incremento de 1 unidade no `S` representa o avanço de uma semana. Por
  exemplo, `S24` é a semana seguinte a `S23`.
- O ficheiro final gerado é referente à semana indicada no ficheiro
  `Simulação PVP`.
- Deve existir um ficheiro da semana anterior, que corresponde ao ficheiro final
  construído na semana anterior.
- Os valores da coluna `Comentarios (face ao suivi)` desse ficheiro anterior
  devem ser copiados para a coluna `HISTORICO` do novo ficheiro com lógica
  equivalente a `PROCX`: procurar primeiro por `ITM8` e, se não existir
  correspondência, procurar por `EAN`.
- As colunas devem ser identificadas pelo cabeçalho, porque a letra pode mudar
  quando o template tem mais ou menos colunas.

## Colunas da Bia

- As colunas `PVP PROMO?` e `MODA?` não fazem parte do template e não devem ser
  adicionadas ao ficheiro final.
- A coluna `pvp folheto bloqueado` deve permanecer na coluna AE, com a fórmula:

```excel
=IF(AC4="","",IF(AC4>=(TODAY()+15),"não","sim"))
```

## Ordenação

O ficheiro final deve ser ordenado por:

1. `Grupo` - coluna C
2. `FIL` - coluna E
3. `FAM` - coluna F
4. `SFA` - coluna G
5. `MARCA` - coluna J
6. `PVP Cadencier` - coluna R
