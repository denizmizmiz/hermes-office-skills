# AC-B03 labeling worksheet (generated aid)

For each row decide the CORRECT semantic role and write it into
your `labels.json` (`expected_role`). The `machine` column is only
the current algorithm's guess - if you copy it blindly the oracle
grades itself and proves nothing.

Allowed roles: identifier, label, date, datetime, period, quantity,
currency, percent, code, address, person_name, free_text, flag,
summary, formula_derived, empty_slot, unknown.

| book | col | header | data_type | format | sample values | machine guess | level |
|---|---|---|---|---|---|---|---|
| clean | A | Ad | string | General | — | label | MEDIUM |
| clean | B | Adet | integer | General | — | quantity | MEDIUM |
| clean | C | Birim Fiyat | decimal | #,##0.00 ₺ | — | currency | HIGH |
| clean | D | Tutar | formula | General | — | currency | HIGH |
| clean | E | Notlar | string | General | — | free_text | HIGH |
| clean | F | Aciklama | string | General | — | label | MEDIUM |
| english | A | Name | string | General | — | label | MEDIUM |
| english | B | Quantity | integer | General | — | quantity | MEDIUM |
| english | C | Unit Price | decimal | #,##0.00 ₺ | — | currency | HIGH |
| english | D | Total | formula | General | — | formula_derived | HIGH |
| english | E | Notes | string | General | — | free_text | HIGH |
| english | F | Comment | string | General | — | free_text | MEDIUM |
| reordered | A | Adet | integer | General | — | quantity | MEDIUM |
| reordered | B | Ad | string | General | — | label | MEDIUM |
| reordered | C | Birim Fiyat | decimal | #,##0.00 ₺ | — | currency | HIGH |
| reordered | D | Tutar | formula | General | — | currency | HIGH |
| reordered | E | Notlar | string | General | — | free_text | HIGH |
| reordered | F | Aciklama | string | General | — | label | MEDIUM |
| mismatch | A | Ad | string | General | — | label | MEDIUM |
| mismatch | B | Adet | integer | General | — | quantity | MEDIUM |
| mismatch | C | Birim Fiyat | decimal | #,##0.00 ₺ | — | currency | HIGH |
| mismatch | D | Tutar | string | General | — | currency | HIGH |
| mismatch | E | Notlar | string | General | — | free_text | HIGH |
| mismatch | F | Aciklama | string | General | — | label | MEDIUM |
| extra | A | Ad | string | General | — | label | MEDIUM |
| extra | B | Adet | integer | General | — | quantity | MEDIUM |
| extra | C | Birim Fiyat | decimal | #,##0.00 ₺ | — | currency | HIGH |
| extra | D | Tutar | formula | General | — | currency | HIGH |
| extra | E | Notlar | string | General | — | free_text | HIGH |
| extra | F | Aciklama | string | General | — | label | MEDIUM |
| extra | G | Fazla Kolon | string | General | — | identifier | LOW |
| missing | A | Ad | string | General | — | label | MEDIUM |
| missing | B | Adet | integer | General | — | quantity | MEDIUM |
| missing | C | Birim Fiyat | decimal | #,##0.00 ₺ | — | currency | HIGH |
| missing | D | Bilinmeyen Alan | string | General | — | label | LOW |
| missing | E | Notlar | string | General | — | free_text | HIGH |
| missing | F | Aciklama | string | General | — | label | MEDIUM |
| duplicate | A | Ad | string | General | — | label | MEDIUM |
| duplicate | B | Adet | integer | General | — | quantity | MEDIUM |
| duplicate | C | Ad | mixed | #,##0.00 ₺ | — | label | LOW |
| duplicate | D | Tutar | formula | General | — | currency | HIGH |
| duplicate | E | Notlar | string | General | — | free_text | HIGH |
| duplicate | F | Aciklama | string | General | — | label | MEDIUM |
| hidden | A | Ad | string | General | — | label | MEDIUM |
| hidden | B | Adet | integer | General | — | quantity | MEDIUM |
| hidden | C | Birim Fiyat | decimal | #,##0.00 ₺ | — | currency | HIGH |
| hidden | D | Tutar | formula | General | — | currency | HIGH |
| hidden | E | Notlar | string | General | — | free_text | HIGH |
| hidden | F | Aciklama | string | General | — | label | MEDIUM |
| ambiguous | A | Ad | string | General | — | label | MEDIUM |
| ambiguous | B | Adet | integer | General | — | quantity | MEDIUM |
| ambiguous | C | Birim Fiyat | decimal | #,##0.00 ₺ | — | currency | HIGH |
| ambiguous | D | Birim Fiyat | integer | #,##0.00 ₺ | — | currency | HIGH |
| ambiguous | E | Notlar | string | General | — | free_text | HIGH |
| ambiguous | F | Aciklama | string | General | — | label | MEDIUM |
| merged | A | Ad | string | General | — | label | MEDIUM |
| merged | B | Adet | integer | General | — | quantity | MEDIUM |
| merged | C | Birim Fiyat | decimal | #,##0.00 ₺ | — | currency | HIGH |
| merged | D | Tutar | formula | General | — | currency | HIGH |
| merged | E | Notlar | string | General | — | free_text | HIGH |
| merged | F | Aciklama | string | General | — | label | MEDIUM |
