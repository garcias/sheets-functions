# Inner Join and Left Join in Google Sheets

During my research to develop `INNERJOIN` and `LEFTJOIN`, I came across a few relevant questions on Stack Overflow. It was a relief to know that I wasn't the only person wanting to do joins in Google Sheets.

## Question: INNER JOIN in Google Spreadsheets

[https://stackoverflow.com/q/65809201](https://stackoverflow.com/q/65809201) by user [White_King](https://stackoverflow.com/users/1180993)

I don't know or remember the technical name of what I'm looking for but I think an example will be enough for you to understand exactly what I'm looking for.

Given table A
```
a   x1
b   x2
c   x1
```

and Table B
```
x1  x
x1  y
x1  z
x2  p
x2  z
```
I Need Table C
```
a   x
a   y
a   z
b   p
b   z
c   x
c   y
c   z
```

I'm looking for a formula or a set of them to get table C

I guess just need to add an extra row on the C table with each value of the first column on the table A for each corresponding value of TableA!Column2 to TableB!Column1 But I can't find how

I think this is a simple SQL Inner Join

## Answers

A couple people suggested Apps Script. [player0](https://stackoverflow.com/users/5632629) did come up with a formula, though it relied on string hacking and would not generalize to different-size tables. The OP also clarified that the solution had to work for non-unique keys. So I decided to add in an alternative solution based on `INNERJOIN`.

### Solution using modern array-based formulas, generalizes to any-size tables

Coming to this question after the advent of array-based formulas like `MAP`, `BYROW`, `LAMBDA`, etc., which seem to be faster than Apps Script functions (when less than 10<sup>6</sup> cells are involved). I want to offer an alternative solution that uses only formulas, and does not require "string hacking" (1), because some people need such features. This solution will work on tables with different shapes.

**Definitions.** In your example, we'll assume Table A is `TableA!A1:B3` and Table B is `TableB!A1:B5`, and we're going to use `LET` to define four variables for clarity:

- `data_left` is `TableA!A1:A3`, represents the **data** from Table A to display
- `keys_left` is `TableA!B1:B3`, represents the **keys** of the data in Table A, that we'll use for matching
- `data_right` is `TableB!B1:B5`, represents Table B data
- `keys_right` is `TableA!A1:A5`, represents Table B keys

In either table, the key values are not unique. Our goal is to find all matches of key values (e.g. `x1` and `x2`) between the two tables, and display only the corresponding values from `TableA!A1:A3` and `TableB!B1:B5`.

**Formula.** The formula below generates a new array containing values from the first column of table A and second column of table B. Each row represents a match between `keys_left` and `keys_right`, with proper duplication when a key appears in multiple rows.

```
= LET(
  data_left, TableA!A1:A3, data_right, TableB!B1:B5,
  keys_left, TableA!B1:B3, keys_right, TableB!A1:A5,
  index_left, SEQUENCE( ROWS( keys_left ) ),
  prefilter, ARRAYFORMULA( MATCH( keys_left, keys_right, 0 ) ),
  index_left_filtered, FILTER( index_left, prefilter ),
  keys_left_filtered, FILTER( keys_left, prefilter ),
  matches, MAP( index_left_filtered, keys_left_filtered, LAMBDA( id_left, key_left,
    LET(
      row_left, XLOOKUP( id_left, index_left, data_left ),
      matches_right, FILTER( data_right, keys_right = key_left ),
      TOROW( BYROW( matches_right, LAMBDA( row_right,
        HSTACK( row_left, row_right )
      ) ) )
    )
  ) ),
  wrapped, WRAPROWS( FLATTEN(matches), COLUMNS(data_right) + COLUMNS(data_left) ),
  notblank, FILTER( wrapped, NOT(ISBLANK(CHOOSECOLS(wrapped, 1))) ),
  notblank
)
```

**How it works?** A few tricks are necessary to make this both accurate and fast:

- **`index_left`**: Create a temporary, primary-key array to index the left table, so that you can retrieve rows from it later.
- **`prefilter`**: Prefilter this index to omit rows with unmatched keys. Use this to filter the index (`index_left_filtered`) and the keys (`keys_left_filtered`) accordingly. (2)
- **`matches`**: For each remaining value in the index:
    - **`row_left`**: `XLOOKUP` the corresponding row from `data_left`
    - **`matches_right`**: use `FILTER` to find all matching rows in the `data_right`
    - concatenate a copy of `row_left` with each matching row from the `data_right`
    - Use `TOROW( BYROW() )` to flatten the resulting array into a single row, because `MAP` can return 1D array for each value of the index but not 2D array. (This makes a mess but we fix it later.)
- **`wrapped`**: The resulting array will have as many rows as the filtered index, but number of columns will vary depending on the maximum number of matches for any given index. Use `WRAPROWS` to properly stack and align matching rows. This leads to a bunch of empty blank cells but ...
- **`notblank`**: ... those are easy to filter out.

**Generalize.** To apply this formula to other tables, just specify the desired ranges (or the results of `ARRAYFORMULA` or `QUERY` operations) for the first four variables; `keys_left` and `keys_right` must be single column but `data_left` and `data_right` can be multi-column. (Or create a Named Function and specify the four variables as parameters as "Argument placeholders".) 

**Named Function.** If you just want to use this, you can import the Named Function `INNERJOIN` from my spreadsheet [functions](https://docs.google.com/spreadsheets/d/1uKanNWKZL3UArI1A14LXNM86qqZTSVQ42ngvg9emCjY/). That version assumes the first row contains column headers. [See documentation at this GitHub repo.](https://github.com/garcias/sheets-functions)

**Notes.**

(1) I loved string-hacking approaches back when they were the only option, but [doubleunary](https://stackoverflow.com/users/13045193) pointed out that they [convert numeric types to strings and cause undesirable side effects](https://stackoverflow.com/a/76126924).

(2) This is counterintuitive because it means you search the `keys_right` twice overall; but I found in testing that if you include unmatched rows in the joining step, is much costlier.
