=UNIQUE(queryData!D:D)

=ARRAYFORMULA(IF(A1:A="", "", VLOOKUP(A1:A, queryData!D:E, 2, FALSE)))

=ARRAYFORMULA(TEXTJOIN(", ", TRUE, UNIQUE(FILTER(queryData!F:F, queryData!D:D=A1))))
