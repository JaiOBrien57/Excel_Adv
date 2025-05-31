# Excel Formula List

## Unpivot Table Formula

### Description

The formula takes the input of a table where the goal is to unpivot the columns of the table into a vertical list of data which can be used for formulas and pivot tables.

### Formula (Lambda Named Format)

```
=LAMBDA(data,
  LET(
    nRows, ROWS(data) - 1,
    nCols, COLUMNS(data) - 1,
    ids, INDEX(data, SEQUENCE(nRows * nCols,,1,1) + 1, 1),
    months, INDEX(data, 1, SEQUENCE(nRows * nCols,,2,1)),
    values, INDEX(data, 
                  INT((SEQUENCE(nRows * nCols,,1,1) - 1)/nCols) + 2, 
                  MOD(SEQUENCE(nRows * nCols,,0,1), nCols) + 2),
    HSTACK(ids, months, values)
  )
)
```
| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Row 1    | Data     | More     |
| Row 2    | Info     | Values   |






text
