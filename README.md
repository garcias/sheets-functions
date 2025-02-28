# sheets-functions
Google Sheets named functions that I find really handy for managing data

I love Google Apps Script, but I prefer Named Functions when possible to implement custom operations in Sheets. In my anecdotal experience, even the hackiest combination of nested `ARRAYFORMULA`s and `QUERY`s will be more performant than a well-designed Apps Script function; and won't present uncertainities or confusion with authorization. I store all my favorite Named functions in a Google Sheet named "functions" and then import functions in new sheets as needed.

## Definiting each Named Function

I tend to follow JavasScript-like conventions (e.g., indentation and bracket alignment) when writing a Named Function. Therefore I define Named Functions in files with `.js` extension. In addition, I use JS function-definition syntax to indicate name of the function, its parameters, and its logic. Consider the example below for the function `XLOOKUPIFERROR`.

```js
// XLOOKUPIFERROR on range of keys, with header column
function XLOOKUPIFERRORARRAY( keyrange, lookup_range, result_range, header, headerrow ) {
  ARRAYFORMULA(
    if(
      row(keyrange)=headerrow,
      header,
      iferror( xlookup( keyrange, lookup_range, result_range ), "" )
    )
  )
}
```

To create this Named Function in the Google Sheets interface:

- In **Named function details**: Set the name as `XLOOKUPIFERROR` and enter a brief description
- In **Argument placeholders**: Create parameters named `keyrange`, `lookup_range`, `result_range`, `header`, `headerrow`
- In **Formula definition**: enter `= ARRAYFORMULA( if( ... ) )`
- In **Additional details**: For each parameter, enter a description of it and provide an example.

- 
