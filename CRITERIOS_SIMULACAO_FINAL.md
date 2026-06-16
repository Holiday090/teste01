# Critérios da simulação final

## Fonte principal

- A informação base do ficheiro final deve seguir a tabela dinâmica `TD Psyco` da
  `S24 - Simulação PVP S24-2026.xlsm`.
- A `TD Psyco` pode estar filtrada no Excel. Para gerar o ficheiro final, devem
  ser considerados todos os dados, como se os filtros da pivot fossem removidos.
- Como a pivot pode estar guardada com cache filtrada, o gerador reconstrói a
  `TD Psyco` sem filtros a partir da folha `Dados`.
- A reconstrução exclui `MERCADONA`, porque a `TD Psyco`/ficheiro final só usa
  as colunas Shopping `CONTINENTE`, `LIDL` e `PINGO-DOCE`.

## Estatuto

- A coluna `ST` do ficheiro final vem sempre da folha `Dados`.

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
