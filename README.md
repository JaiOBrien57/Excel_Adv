# Excel_Adv


````
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
)`
