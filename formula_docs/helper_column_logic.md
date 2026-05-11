# Formula Documentation

## 1. Helper Column Formula (Raw_Data Sheet — Column A)

```excel
=IF(Search_Bar!$D$8="", 0,
    IF(CHOOSE(Search_Bar!$M$1,
        LEFT([@[Customer Name]], LEN(Search_Bar!$D$8))=Search_Bar!$D$8,
        RIGHT([@[Customer Name]], LEN(Search_Bar!$D$8))=Search_Bar!$D$8,
        ISNUMBER(SEARCH(Search_Bar!$D$8, [@[Customer Name]]))
    ), MAX($A$1:A1)+1, 0)
)
```

### How it works:
- If search box (D8) is empty → returns 0, nothing shows
- CHOOSE reads radio button value from M1 → picks logic branch
  - M1 = 1 → LEFT match (starts with)
  - M1 = 2 → RIGHT match (ends with)
  - M1 = 3 → SEARCH match (contains anywhere)
- If row matches → gets MAX+1 (sequential rank)
- If no match → returns 0

---

## 2. S.No Formula (Search_Bar Sheet)

```excel
=IF($D$8="","",IF(ROW()-ROW($A$11)>MAX(Raw_Data!$A:$A),"",ROW()-ROW($A$11)))
```

### How it works:
- If search box empty → shows nothing
- Calculates sequential row numbers (1,2,3...) for output
- Stops when row number exceeds total matched results
- Prevents blank rows from showing after results end

---

## 3. Customer Name Formula (Search_Bar Sheet)

```excel
=IF($D$8="","",IFERROR(VLOOKUP($A12,Data1,COLUMNS(Search_Bar!$A$10:B11),0),""))
```

### How it works:
- If search box empty → shows nothing
- VLOOKUP looks up S.No (A12) in Data1 named range
- COLUMNS() dynamically picks correct column — no hardcoding
- IFERROR handles unmatched rows silently — no error shown

---

## Logic Flow Summary
 | Step | Action |
|------|--------|
| 1 | User types in Search Box (D8) |
| 2 | Radio Button value in M1 → 1, 2, or 3 |
| 3 | Helper Column assigns rank to matching rows |
| 4 | S.No column generates 1, 2, 3... |
| 5 | VLOOKUP pulls correct data for each rank |
| 6 | Clean filtered output — real time ✅ |
---

## Compatibility
Works on Excel 2007, 2010, 2013, 2016, 2019, 365
No VBA. No macros. No array formulas.
