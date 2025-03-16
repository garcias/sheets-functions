# Stack Exchange posts

I learned a lot of formula tricks through StackOverflow, so I should try to give back when possible.


## How do I cross join two, multi-column tables in Google Sheets — without string hacking or Apps Script?

I posted the following question to StackOverflow, and made up a toy example. 

(Although I implied I was looking for a generalizable solution, I learned that most responders will focus on the example, so in the future I'll need to specify that as a criterion of the desired answer. I ended up editing such answers to generalize them, or posting my own answers to build upon a more specific answer.)

### Question: How do I cross join two, multi-column tables in Google Sheets — without string hacking or Apps Script?

[https://stackoverflow.com/q/79505084](https://stackoverflow.com/q/79505084)

*(I will post my own answer, but I want to pose the question for the sake of others in a similar situation; and to encourage alternative answers.)*

Sometimes I have two tables and I need to cross join them, i.e., generate the combinations of each row of the first table with each row of the second table. For example, suppose I have one table of chefs and the dishes they plated for a competition, and a table of evaluators and the attributes that each specializes on. (I'll specify them as formulas for easy pasting into Google Sheets.)

```
table1 = {
  { "Albion"    , "Artichoke Soufflé Omelett" };
  { "Burgess"   , "Lemony Braised Chicken" };
  { "Hamad"     , "Mabo Dofu Smoothie" };
  { "Berengari" , "Chicken-Fried Plantains" };
  { "Sengupta"  , "Smoky Vegan Corn Salad" }
}

table2 = {
  { "Cho"       , "flavor" };
  { "Nikkelson" , "texture" };
  { "Rodríguez" , "process" }
}
```

Then the cross join of these two tables has 15 rows beginning with:
```
output = {
  { "Albion"    , "Artichoke Soufflé Omelett" , "Cho"       , "flavor" }
  { "Albion"    , "Artichoke Soufflé Omelett" , "Nikkelson" , "texture" }
  { "Albion"    , "Artichoke Soufflé Omelett" , "Rodríguez" , "process" }
  { "Burgess"   , "Lemony Braised Chicken"    , "Cho"       , "flavor" }
  { "Burgess"   , "Lemony Braised Chicken"    , "Nikkelson" , "texture" }
  { "Burgess"   , "Lemony Braised Chicken"    , "Rodríguez" , "process" }
  ...
}
```

I am looking for a formula or Named Function to do this — not Google Apps Script. In addition, I want to avoid "string-hacking" methods in which you serialize arrays into strings with weird delimiters, then do text manipulation and concatenation that generates the desired structure, and finally deserialize into the reshaped array. As [doubleunary](https://stackoverflow.com/users/13045193/doubleunary) pointed out in [their answer](https://stackoverflow.com/a/76126924) to [a similar question](https://stackoverflow.com/questions/42805885/generate-all-possible-combinations-for-columnscross-join-or-cartesian-product), such methods have side effects when they convert numeric types to strings. Personally I found they can have unpredictable behavior if any of the data contain certain emoji. I also find them difficult to troubleshoot and maintain.

My question is similar to the three below, but they were only about cross joining *single-column* tables.
- [Generate all possible combinations for Columns(cross join or Cartesian product)](https://stackoverflow.com/questions/42805885/generate-all-possible-combinations-for-columnscross-join-or-cartesian-product)
- [How to cross join 2 lists?](https://stackoverflow.com/questions/70440464/how-to-cross-join-2-lists)
- [Google sheets - cross join / cartesian join from two separate columns](https://stackoverflow.com/questions/65766369/google-sheets-cross-join-cartesian-join-from-two-separate-columns)

There is a previous question that doesn't ask about multi-column tables explicitly, but includes one in its sample data; but the OP of the question allowed Apps Script: 
- [How to perform Cartesian Join with Google Scripts & Google Sheets?](https://stackoverflow.com/questions/65556675/how-to-perform-cartesian-join-with-google-scripts-google-sheets)

This question *might* be similar, but I can't tell because the OP did not include sample data in the question and the linked spreadsheet no longer exists.
- [Google Sheets Cross Join Function Tables with More than Two Columns](https://stackoverflow.com/questions/60572866/google-sheets-cross-join-function-tables-with-more-than-two-columns)


### Answers

Almost immediately, a couple people proposed clever solutions, and I learned something new from each of them.

#### PatrickdC's answer

[PatrickdC's answer](https://stackoverflow.com/a/79505103) used `HSTACK` on each combination of rows from each table to accomplish the horizontal concatenation. (My method does concatenation at the end.) 

```
= LET(
  table1, A2:B6, table2, E2:F4,
  WRAPROWS(
    TOROW( BYROW( table1, LAMBDA( x,
      TOROW( BYROW( table2, LAMBDA( y,
        HSTACK( x, y)
      ) ) ) 
    ) ) ),
    COLUMNS( table1 ) + COLUMNS( table2 )
  )
)
```

They applied `BYROW` to each table, which surprised me because `BYROW` tends to be finicky about its lambda returning an array. Their trick was to wrap it in `TOROW`. For example, `BYROW( table2, LAMBDA( y, HSTACK( x, y) ) )`, produces a 2D array, which cannot be returned to `BYROW( table1, LAMBDA( x, ... )`, e.g.

```
{ { chef1, dish2, judge1, speciality1 };
  { chef1, dish1, judge2, speciality2 };
  ... }
```

So `TOROW` flattens that into a single-row array:

```
{ chef1, dish1, judge1, speciality1, chef1, dish1, judge2, speciality2, ... }
...
```

Which *can* be returned to the outer `BYROW` loop. This ultimately leads to a 2D array with the wrong shape:

```
{ { chef1, dish1, judge1, speciality1, chef1, dish1, judge2, speciality2, ... };
  { chef2, dish2, judge1, speciality1, chef2, dish2, judge2, speciality2, ... };
  ... }
```

But then `WRAPROWS` reshapes this to the desired result

```
{ { chef1, dish1, judge1, speciality1 };
  { chef1, dish1, judge2, speciality2 };
    ...
  { chef2, dish2, judge1, speciality1 }; 
  { chef2, dish2, judge2, speciality2 }; 
    ...;
  ... }
```
