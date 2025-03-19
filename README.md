# sheets-functions
Google Sheets named functions that I find really handy for managing data

I love Google Apps Script, but I prefer Named Functions when possible to implement custom operations in Sheets. In my anecdotal experience, even the hackiest combination of nested `ARRAYFORMULA`s and `QUERY`s will be more performant than a well-designed Apps Script function; and won't present uncertainities or confusion with authorization. I store all my favorite Named functions in a Google Sheet named "functions" and then import functions in new sheets as needed.

Current version of spreadsheet with all my current functions already entered: [functions](https://docs.google.com/spreadsheets/d/1uKanNWKZL3UArI1A14LXNM86qqZTSVQ42ngvg9emCjY/template/preview)

## Featured functions

- `CROSSJOIN`. Cartesian product of two tables. 
- `MELT`. Unpivot selected columns of a wide table, creating a "tall and thin" version with the same data.
- `INNERJOIN`. Join two tables on keys, even if key values are not unique in either table.
- `QBN`. Like QUERY, but you refer to each column by its heading instead of "A", "B", etc. Developed by Stack Exchange user [carecki](https://webapps.stackexchange.com/questions/57540/can-i-use-column-headers-in-a-query/167714#167714).

## Specify each Named Function in a consistent way

I tend to follow JavasScript-like conventions (e.g., indentation and bracket alignment) when writing a Named Function. Therefore I define them in this repo as files with `.js` extensions, and I use JSDoc-like syntax to indicate the name of the function and to describe its parameters. Different parts of the definition need to be translated to the Named Functions interface in Sheets. Consider the example below for the function `XLOOKUPIFERRORARRAY`.

```js
/**
 * XLOOKUPIFERRORARRAY
 *   Applies XLOOKUPIFERROR to a range of keys, with header column
 * 
 * @param {Range} keyrange - range of keys, e.g. A:A
 * @param {Range} lookup_range - range to search for each key, e.g. sheet1!D:D
 * @param {Range} result_range - range that contains result values, e.g. sheet1!A:A
 * @param {Text} header - text to display as the column header, e.g. "Last name"
 * @param {Integer} headerrow - row number of header, e.g. 1
 * @return {Array} an Array of results, one for each value in the keyrange
 */

= ARRAYFORMULA(
  IF(
    row(keyrange)=headerrow,
    header,
    IFERROR( XLOOKUP( keyrange, lookup_range, result_range ), "" )
  )
)
```

To create this Named Function in the Google Sheets user interface:

- In **Named function details**: Set the name as `XLOOKUPIFERRORARRAY` and enter the description *"XLOOKUPIFERROR on a range of keys, with header column"*
- In **Argument placeholders**: Create parameters named `keyrange`, `lookup_range`, `result_range`, `header`, `headerrow`
- In **Formula definition**: Enter `= ARRAYFORMULA( if( row(keyrange)=headerrow, ... ) )`
- In **Additional details**: For `keyrange` enter the description *"range of keyes"* and the example *"A:A"*. Repeat for remaining arguments


## Change log

### 2025-03-18: Update INNERJOIN for 2X speed improvement

Overall algorithm still the same, I just found three slight changes would enhance performance by a factor of 2X:

- Use `SEQUENCE` to generate the index for left table.
- Use `MATCH` to more quickly identify left keys that exist in right keys, and thus build the prefilter.
- Prefilter both the left index and left keys, and pass both to `MAP` to allow for parallel iteration (and thus saves us from a costly lookup of each key value).

It's too bad we couldn't also pass the entire left table for parallel iteration, but unfortunately `MAP` sees that multi-column array as different size from the key column array, even though they're the same height.

### 2025-03-16: Update testing functions to be array based

I updated the testing functions so they generate arrays of sample values at a time, because most of my testing is on array formulas. Each can still generate a single random value by setting the desired dimensions as 1 row, 1 column. They are named `TESTALPHAARRAY`, `TESTASCIIARRAY`, `TESTSELECTARRAY`, `TESTSELECTMULTIPLEARRAY`.

### 2025-03-07: New formula INNERJOIN

On a roll! And I finally achieved a holy grail of Google Sheets array manipulation: a generic, *symmetrical*, *performant* implementation of an inner join on *foreign* keys. It will correctly duplicate rows from either the left or right table as needed, and will omit rows where no match is found on the key. The implementation is surprisingly fast, I don't see lagging until I get to computing 10^7 cells (`ARRAY_CONSTRAIN`ed). Performance is best when the left table is shorter than the right table.

Of course this isn't as nice as having joins defined and optimized in the Google Visualization Query language, this is just for relatively simple tasks, where you join two tables first and then wrap it in a `QUERY` formula to get more precisely what you want. If you need more complex queries in Google Sheets, and you're able to use Apps Script functions in your situation, then check out the [gsSQL project](https://github.com/demmings/gsSQL) at https://github.com/demmings/gsSQL.

My next goal is to implement a left join, for cases where I need to see unmatched rows in a given table. I suspect this will be a little slower, but we'll see ...

A few tricks were necessary to make `INNERJOIN` work accurately and quickly:

- Create a temporary, primary-key array to index the left table, so that you can retrieve rows from it later
- Prefilter this index to omit rows with unmatched keys. This is counterintutitive because it means you search the right table twice overall; but I found that including unmatched rows in the joining step is much costlier.
- For each remaining value in the index:
    - Look up the corresponding left-table key, and use `FILTER` to find *potentially multiple* matching rows in the right table as an array. 
    - Then `XLOOKUP` the corresponding row from the left table, and concatenate a copy of that row with each matching row from the right table. 
    - Finally `FLATTEN` those rows since `MAP` can return 1D array for each row but not 2D array. (This makes a mess but we fix it later.)
- The resulting array will have as many rows as the filtered index, but number of columns will vary depending on the maximum number of matches for any given index. Use `WRAPROWS` to properly stack matching rows. This leads to a bunch of empty blank cells but those are easy to filter out.

**TIL.** `XLOOKUP` can return an array, if you specify the `result_range` as a multiple-column array! (But this behavior is unpredictable if you put `XLOOKUP` into an `ARRAYFORMULA`, which limits its usefulness.)

### 2025-03-07: Update MELT

I learned a lot from rewriting `CROSSJOIN` that looked immediately applicable to `MELT`. The new version is much more performant, and runs 4-6X faster than the previous version. I was able to melt a 10^5 row x 10 column array with just a slight lag (the result was `ARRAY_CONSTRAINED`, or else it would need more rows to be created to accommodate the output). It's difficult to test beyond that because creating 10^6 rows makes the sheet laggy all on its own, even without computing formulas.

Out of nostalgia and respect, I kept the old version and renamed it `MELT_LEGACY`. I think it's important to acknowledge and remember the hacky but ingenious strategies we used to rely on, and which enabled people to do amazing things despite the enterprise environments that restricted use of SQL, jq, Python, and other tools. I also wanted to retain a memory of how much people relied on each other for ideas, like the brilliant header-concatenation trick I learned from Prashanth KV over at Info Inspired.

### 2025-03-05: Update CROSSJOIN

The previous version of `CROSSJOIN` relied heavily on a type of strategy I call "serialize - manipulate - deserialize" (SMD). In an SMD-type approach, you:

- **Serialize:** use `JOIN` or `TEXTJOIN` to convert an array of values into a single string (or array of strings) of those values separated by an unusual delimiter
- **Manipulate:** use `REPT`, `&`, `FLATTEN`, `TRANSPOSE`, and regex functions to represent the rearrangement of values in the desired transformation
- **Deserialize:** use `SPLIT` to convert the manipulated strings into structured array values

SMD formulas were necessary for complex array transformations because of the limited set of array functions in Google Sheets. They were ingenious hacks that made Sheets incredibly useful for a greater range of data analysis and management tasks, and enabled people to be much more imaginative about what you can do in a spreadsheet. But they have disadvantages that feel quite limiting today: (a) they are very difficult to read and thus to troubleshoot, (b) in some situations they can run into character limits on strings, (c) they can easily get confused by empty cells or unusual characters in the source data.

Today we have advanced features like `BYROW`, `MAKEARRAY`, `MAP`, `LAMBDA`, etc. that can accomplish the same transformations but don't have these limitations and are way easier to troubleshoot. So as fond as I am of SMD formulas, I want to rewrite my array-transformation functions to use the new language features. 

The new version of `CROSSJOIN` is my first attempt at this process, and I'm pretty happy with it. It nicely shows the symmetry of the Cartesian product, and it seems just as performant as the SMD-based version (the sheet will start to lag if the result of this function exceeds 10^5 rows, even if wrapped with `ARRAY_CONSTRAIN`). In addition, the new version can take multiple-column arrays and will compute the Cartesian product between each row properly; though of course this will reduce the number of rows you can generate before you notice lagging.

In the process, I learned to use `FLATTEN`, `WRAPROWS`, and `WRAPCOLS` to great effect in reshaping arrays. `BYROW` isn't as useful as I thought for duplicating rows or columns; it's difficult to get it to accept an array returned by its lambda function. But I learned that `MAKEARRAY` handles returned arrays more easily, even if you specify creation of a single-row array, as long as each array is orthogonal. This is crucial to create duplicates each row or column in a source array, which is necessary for the Cartesian product; and I suspect this will be useful to rewrite `MELT` as well.
