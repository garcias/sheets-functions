### Edit the question

**Update.** Here is a markdown representation of the input.

|        | A                              |     B    |     C    |     D    |     E    |
| ------:|:------------------------------ |:-------- |:-------- |:-------- |:-------- |
| **1**  | *Email*                        | *Data 1* | *Data 2* | *Data 3* | *Data 4* |
| **2**  | email@email.com                |  TRUE    |  TRUE    |  TRUE    |  TRUE    |
| **3**  | me@other.email.com             |  FALSE   |  TRUE    |  TRUE    |  TRUE    |
| **4**  | someone@this.email.com         |  TRUE    |  TRUE    |  TRUE    |  TRUE    |
| **5**  | yesno@again.com                |  FALSE   |  FALSE   |  FALSE   |  FALSE   |
| **6**  | memyselfandi@info.org          |  TRUE    |  TRUE    |  TRUE    |  TRUE    |
| **7**  | youyourselfandthou@english.org |  FALSE   |  FALSE   |  TRUE    |  TRUE    |
| **8**  | foultarnished@omen.king.gov    |  FALSE   |  TRUE    |  FALSE   |  TRUE    |
| **9**  |                                |          |          |          |          |
| **10** | USE                            |  TRUE    |  TRUE    |  FALSE   |  TRUE    |
| **11** | INVERT                         |  TRUE    |  FALSE   |  FALSE   |  FALSE   |

(The edit wasn't accepted.)

### Formula

The following formula uses `LET` to define two ranges according to your example. The **selections** range includes both the column of emails and the multiple columns of checkboxes; the **options** range includes the two rows of options to use a column's selections and to invert those selections. The formula makes positional assumptions: that the first column of **selections** is emails and first column of **options** is the "use" and "invert" labels. It assumes the "use" row is always above the "invert" row.

The two ranges must be same width, and you can extend these ranges to include any new columns of checkboxes you create. Alternatively, you can set "open-ended" references, e.g., `selections, A2:8, options, A10:11` will include all columns that exist in the sheet; but you must be sure that all columns to the right contain no other data.

```
= LET(
  selections, A2:E8, options, A10:E11,
  remove_first_column, LAMBDA( array, CHOOSECOLS( array, SEQUENCE( 1, COLUMNS( array ) - 1, 2 ) ) ),
  emails,  CHOOSECOLS( selections, 1 ),
  selects, remove_first_column( selections ),
  use,    CHOOSEROWS( remove_first_column( options ), 1 ),
  invert, CHOOSEROWS( remove_first_column( options ), 2 ),
  apply_options, LAMBDA( select, use, invert, 
    IF( use, XOR( select, invert ), TRUE )
  ),
  includes, BYROW( selects, LAMBDA( select,
    MAP( select, use, invert, LAMBDA( select, use, invert,
      apply_options( select, use, invert ) 
    ) )
  ) ),
  all, BYROW( includes, LAMBDA( select, AND( select ) ) ),
  selected_emails, IFERROR( FILTER( emails, all ) ),
  csv, TEXTJOIN( ",", , selected_emails ),
  csv
)
```

### How it works

- Set four arrays **`emails`**, **`selects`**, **`use`**, **`invert`** to represent each block of information in the input ranges. 
- **`apply_options`**: Define a lambda function to compute the following logic: if `use` is true, then return either the value of the selection (if `invert` is false) or its inverse (if `invert` is true); otherwise return true.
- **`includes`**: Iterate through the rows of `selects`. For each row `select`:
    - `select` is an array of boolean values. For each value in that array, get the corresponding values of `use` and of `invert` and compute `apply_options`. This will result in an array of boolean values for each row.
- **`all`**: Iterate through the rows of `includes`. Using `AND` on an array of booleans will return true only if every value is true. (This behavior of `AND` is analogous to Python `all` function.) The result is a single-column array of boolean values, one for each email. 
- **`selected_emails`**: Filter the array `emails` using values in `includes`.
- **`csv`**: Convert the array `selected_emails` to a comma-separated-value format.

If you ever need to troubleshoot this formula, replace `csv` in the last line with `includes`, `all`, or `selected_mails`. 

### Results

Applied to your example, these are the results of each one (if the formula is in `B13`, to help compare with the input).

`includes`

|        | A |    B   |    C   |    D   |    E   |
| ------:|:-:|:------ |:------ |:------ |:------ |
| **13** |   | FALSE  | TRUE   | TRUE   | TRUE   |
| **14** |   | TRUE   | TRUE   | TRUE   | TRUE   |
| **15** |   | FALSE  | TRUE   | TRUE   | TRUE   |
| **16** |   | TRUE   | FALSE  | TRUE   | FALSE  |
| **17** |   | FALSE  | TRUE   | TRUE   | TRUE   |
| **18** |   | TRUE   | FALSE  | TRUE   | TRUE   |
| **19** |   | TRUE   | TRUE   | TRUE   | TRUE   |

`all`

|        | A |    B   |
| ------:|:-:|:------ |
| **13** |   | FALSE  |
| **14** |   | TRUE   |
| **15** |   | FALSE  |
| **16** |   | FALSE  |
| **17** |   | FALSE  |
| **18** |   | FALSE  |
| **19** |   | TRUE   |

`selected_emails`

|        | A | B                              |
| ------:|:-:|:------------------------------ |
| **13** |   | me@other.email.com             |
| **14** |   | foultarnished@omen.king.gov    |
