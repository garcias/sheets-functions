# Left Join in Google Sheets

During my research to develop `INNERJOIN` and `LEFTJOIN`, I came across a few relevant questions on Stack Overflow. It was a relief to know that I wasn't the only person wanting to do joins in Google Sheets.

## Question: google sheets query left join one-to-many

[https://stackoverflow.com/q/65363849](https://stackoverflow.com/q/65363849) by user [Jonathan Livingston Seagull](https://stackoverflow.com/users/1131165)

I have 2 tables and I am trying to perform a left join using google query language,or any formula that could output the result set.

`Table1!A1:A4`

|     |**A**|
|-----|-----|
|**1**| ID  |
|**2**| 1   |
|**3**| 2   |
|**4**| 3   |

`Table2!A1:B5`

|     |**A**     |**B**|
|-----|---------:|-----|
|**1**| Payment  | ID  |
|**2**| 30       | 2   |
|**3**| 20       | 3   |
|**4**| 10       | 3   |
|**5**| 50       | 1   |

`'Result set'!A1:B6`

|     |**A**|  **B**   |
|-----|-----|---------:|
|**1**| ID  | Payment  |
|**2**| 1   | 50       |
|**3**| 2   | 30       |
|**4**| 3   | 20       |
|**5**| 3   | 10       |
|**5**| 4   |          |


## Answers

One person suggested a string-hacking approach, another offered Apps Script.

### Generalized solution using modern array-based formulas

Here is an alternative solution that uses array-based formulas like `MAP`, `BYROW`, `LAMBDA`. Some users may not be able to use Apps Script (due to employer restrictions) or string-hacking methods (due to [side effects on type conversion](https://stackoverflow.com/a/76126924)), so this solution would work for them. It also generalizes to multiple-column tables, i.e., it will combine multiple columns from the two tables.

**Definitions.** In your example, we'll assume Table1 is in a range on one sheet (`TableA!A1:B3`) and Table2 is in a range on another (`TableB!A1:B5`). The desired result will go into the range `'Result set'!A1:B6`. (I edited your question to indicate the assumed ranges of each example table so I can reference them better; I chose range specifiers that are consistent with pre-existing answers.)

**Formula.** The formula below defines a lambda function with four parameters, and then invokes it using the following ranges fromm your exmaple tables as arguments. 

| **parameter** | **argument**   |
|---------------|----------------|
| `data_left`   | `Table1!A2:A4` |
| `keys_left`   | `Table1!A2:A4` |
| `data_right`  | `Table2!A2:A5` |
| `keys_right`  | `Table2!B2:B5` |

```
= LAMBDA(
  data_left, keys_left, data_right, keys_right,
  LET(
    index_left, SEQUENCE( ROWS( keys_left ) ),
    matches, MAP( index_left, keys_left, LAMBDA( id_left, key_left,
      LET(
        row_left, XLOOKUP( id_left, index_left, data_left ),
        matches_right, IFERROR( FILTER( data_right, keys_right = key_left ), ),
        TOROW( BYROW( matches_right, LAMBDA( row_right,
          HSTACK( row_left, row_right )
        ) ) )
      )
    ) ),
    wrapped, WRAPROWS( FLATTEN(matches), COLUMNS(data_right) + COLUMNS(data_left) ),
    notblank, FILTER( wrapped, NOT(ISBLANK(CHOOSECOLS(wrapped, 1))) ),
    notblank
  )
)( `Table1!A2:A4`, `Table1!A2:A4`, `Table2!A2:A5`, `Table2!B2:B5` )
```

**How it works?** 

- **`index_left`**: Create a temporary, primary-key array to index the left table, so you can retrieve rows from it later.
- **`matches`**: For each value in the index:
    - **`row_left`**: `XLOOKUP` the corresponding row from `data_left`
    - **`matches_right`**: use `FILTER` to find all matching rows in the `data_right`; or `NA()` if there were none
    - for each row in `matches_right` (or `NA()` if no matches), concatenate it with a copy of `row_left`
    - Use `TOROW()` to flatten the resulting array into a single row, because `MAP` can return 1D array for each value of the index but not 2D array. (This makes a mess but we fix it later.)
- **`wrapped`**: The resulting array will have as many rows as the filtered index, but number of columns will vary depending on the maximum number of matches for any given key. Use `WRAPROWS` to properly stack and align matching rows. This leads to a bunch of empty blank rows but ...
- **`notblank`**: ... those are easy to filter out.

**Generalize.** To apply this formula to other examples, just specify the desired ranges (or the results of `ARRAYFORMULA` or `QUERY` operations) for the first four variables; `keys_left` and `keys_right` must be single column but `data_left` and `data_right` can be multi-column and the resulting array will contain all columns of both. (If you expect to use this a lot, you could create a Named Function with the four parameters as "Argument placeholders", and with the body of the `LAMBDA` as the "Function definition".)

**Named Function.** If you just want to use this, you can import the Named Function `LEFTJOIN` from my spreadsheet [functions](https://docs.google.com/spreadsheets/d/1uKanNWKZL3UArI1A14LXNM86qqZTSVQ42ngvg9emCjY/). That version assumes the first row contains column headers. [See documentation at this GitHub repo.](https://github.com/garcias/sheets-functions). Screenshot below shows application of this named function in cell I1: `="LEFTJOIN( E:G, F:F, B:C, A:A)"`. Note that numeric-type data in the source tables remain same type in the result.

![Screenshot of a single sheet in Google Sheets. Columns A:C contain entries for the dish recipes, with columns for dish, ingredient, and amount. Columns E:G contain entries for patron orders, with columns for diner, dish, number. Columns I:M show the results of a left join operation, with columns for diner, dish, number, ingredient, amount.](left-join-img.png)
