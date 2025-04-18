# Template literal for string interpolation

I really don't like concatenating values in Google Sheets, it's just so hard to read, especially when constructing text for `QUERY`. So I looked around for string interpolation methods but didn't find any. The closest I came to it was a question on StackOverflow looking for the Google Sheet equivalent of Python f-strings. The only answer suggested `CONCATENATE`, which IMO is not much better than using `&` operator because there are still lots of intervening quotation marks.

I learned the `REDUCE` formula a while ago and was itching for an opportunity to use it, and this seemed like a great use case because it requires iteration, and `REDUCE` is closest thing Sheets comes to an interation mechanism. 


## Question: Is there an equivalent of an f-string in Google Sheets?

[https://stackoverflow.com/q/68178453](https://stackoverflow.com/q/68178453) by user [cryptoversion](https://stackoverflow.com/users/16342128)

I am making a portfolio tracker in Google Sheets and wanted to know if there is a way to link the "TICKER" column with the code in the "PRICE" column that is used to pull JSON data from Coin Gecko. I was wondering if there was an f-string like there is in Python where you can insert a variable into the string itself. Ergo, every time the Ticker column is updated the coin id will be updated within the API request string. Essentially, string interpolation


## Answers

The only answer suggested to use `CONCATENATE`. I wanted to offer something closer to the specific question itself, which was for string interpolation.

### Formula to interpolate values into a template string containing placeholders

The OP asked for string interpolation similar to Python f-string. The goal is that a user only needs to enter a single string to define a template, but within that string there is specially marked "placeholder" to indicate where to substitute a value. The placeholder can be a **key** enclosed in curly braces, e.g. `{TICKER}`. Google Sheets didn't have such a formula.

Fast forward to 2025, Sheets still doesn't have it, but does offer array-based formulas and iteration mechanism like `REDUCE`. Combined with `REGEXREPLACE`, we can define a Named Formula that simulates simple interpolation. 

**Formula.** The formula below takes three parameters:

- *`template`*: a string that contains placeholders enclosed in curly braces
- *`keys`*: a string to define the placeholder key (or single-row array if multiple placeholders)
- *`values`*: a value corresponding to the key (or single-row array if multiple placeholders)

Create a Named Function `TEMPLATE` with these three parameters, so named and in that order, then enter the formula definition:

```
= REDUCE( template, keys, LAMBDA( acc, key,
  LET(
    placeholder, CONCATENATE( "\{", key, "\}" ),
    value, XLOOKUP( key, keys, values ),
    REGEXREPLACE( acc, placeholder, TO_TEXT(value) ) 
  )
 ) )
```

**Example with one placeholder.** Your example uses a custom Apps Script function called ImportJSON, but your question is more about the string interpolation, so I will just focus on how to generate the URL based on the value of A2 (the cell containing string "BTC"). In a cell enter:

```
= TEMPLATE(
    "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&ids={TICKER}",
    "TICKER", A2
)
```

The result should be `https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&ids=BTC`

**Example with multiple placeholders.** You can give arrays as the `keys` and `values` arguments for multiple placeholders:

```
= TEMPLATE( 
  "My name is {name} and I am {age} years old.",
  { "name", "age" }, { "Lan", 67 }
)
```

Confirm the result is: `My name is Lan and I am 67 years old.`

**How it works?** The `REDUCE` formula takes the string `template` as an initial value, and iterates through the values of `keys`. In each iteration:
- It constructs a regular expression to match the placeholder named by the key, e.g. `"\{name\}"`.
- It looks up the value corresponding to the key, e.g. `"Lan"`.
- It uses `REGEXREPLACE` on the template to match all instances of the placeholder and replace each one with the value.
- It passes the resulting string, which may contain other placeholders (e.g. `"My name is Lan and I am {age} years old."`) to the next iteration.

**Named Function.** If you just want to use this, you can import the Named Function `TEMPLATE` from my spreadsheet [functions](https://docs.google.com/spreadsheets/d/1uKanNWKZL3UArI1A14LXNM86qqZTSVQ42ngvg9emCjY/). See the [documentation at this GitHub repo](https://github.com/garcias/sheets-functions) for more details.
