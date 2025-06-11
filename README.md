# Excel Formula List

## Unpivot Table Formula

### Description

The formula takes the input of a table where the goal is to unpivot the columns of the table into a vertical list of data which can be used for formulas and pivot tables.

### Formula (Lambda Named Format)

```
=LAMBDA(table,LET(
    data, INDEX(table, SEQUENCE(ROWS(table)-1, 1, 2), SEQUENCE(1, COLUMNS(table))),
    names, INDEX(data,,1),
    headers, INDEX(table,1,),
    nRows, ROWS(data),
    nCols, COLUMNS(data)-1,
    rowSeq, SEQUENCE(nRows*nCols),
    nameCol, INDEX(names, INT((rowSeq-1)/nCols)+1),
    monthCol, INDEX(headers, 1, MOD(rowSeq-1, nCols)+2),
    valueCol, INDEX(data, INT((rowSeq-1)/nCols)+1, MOD(rowSeq-1, nCols)+2),
    HSTACK(nameCol, monthCol, valueCol)
))
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

## Amex Formula

### Description

Takes a value and adds on the amex value (1% currently, change as needed) as a markup and displays the output in a table.

### Formula (Lambda Named Format)

```
=LAMBDA(amount,HSTACK(VSTACK("Invoice Amount:","Total AMEX:","Fee"),VSTACK(amount,ROUND(amount/(1-0.01),2),ROUND((amount/(1-0.01))-amount,2))))
```
### Usage

**Input:**

$100.00

**Output:**

|----------|--------|------|
| Person 1 | Game 1 | Data |
| Person 1 | Game 2 | Data |
| Person 1 | Game 3 | Data |




