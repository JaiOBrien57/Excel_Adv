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
### Usage

**Input:**

Data it takes in is the entire table including the headers and first column:

|          | Game 1 | Game 2 | Game 3 |
|----------|--------|--------|--------|
| Person 1 |  Data  |  Data  |  Data  |
| Person 2 |  Data  |  Data  |  Data  |
| Person 3 |  Data  |  Data  |  Data  |

**Output:**

Will tranform the data into the following output:

|  People  |  Games | Data |
|----------|--------|------|
| Person 1 | Game 1 | Data |
| Person 1 | Game 2 | Data |
| Person 1 | Game 3 | Data |
| Person 2 | Game 1 | Data |
| Person 2 | Game 2 | Data |
| Person 2 | Game 3 | Data |
| Person 3 | Game 1 | Data |
| Person 3 | Game 2 | Data |
| Person 3 | Game 3 | Data |




