# sheets-functions
Google Sheets named functions that I find really handy for managing data

I love Google Apps Script, but I prefer Named Functions when possible to implement custom operations in Sheets. In my anecdotal experience, even the hackiest combination of nested `ARRAYFORMULA`s and `QUERY`s will be more performant than a well-designed Apps Script function; and won't present uncertainities or confusion with authorization. I store all my favorite Named functions in a Google Sheet named "functions" and then import functions in new sheets as needed.

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
