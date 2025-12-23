
### Formula

Use the following formula in cell E2. I wrote it as a `LAMBDA` function that takes three parameters `id_col`, `start_col`, `end_col`; and is called on the three columns in your example. This makes it easy to convert into a Named Function if you want (see below).

```
= LAMBDA( id_col, start_col, end_col, 
  LET(
    date_sequence, LAMBDA( start, end, LET(
      num_days, DATEDIF( start, end, "D" ) + 1,
      SEQUENCE( num_days, 1, start )
    ) ),
    prepend_id, LAMBDA( id, dates, 
      TOROW( BYROW( dates, LAMBDA( row, { id, row } ) ) )
    ),
    table, MAP( id_col, start_col, end_col, LAMBDA( id, start, end, 
      prepend_id( id, date_sequence( start, end ) )
    ) ),
    WRAPROWS( TOROW(table, 1), 2 )
  )
) ( A2:A4, B2:B4, C2:C4 )
```

### Named function

To convert it into a reusable formula, select the menu items: **Data > Named Functions > Add new function**. Give it the desired name, and then add three "Argument placeholders" witht the exact names `id_col`, `start_col`, `end_col`. Finally for the "Formula definition" copy in all lines of the formula above, except for the first and last lines. Then use the function instead of the formula. For example, if you named the function "EXPAND_DATES" then in E2 you would use `=EXPAND_DATES( A2:A4, B2:B4, C2:C4 )` to get the same result.


### How it works

**`date_sequence`**. Define a function that creates an array of dates, one day apart, from the given start date to the given end date.

**`prepend_id`**. Define a function that takes any given value `id`, and an array of `dates`, and creates an array that alternates between `id` and each date in `dates`. For example, `prepend_id( "A", { 1/20/2025, 1/21/2025 } )` returns a single-row, 4-column array `{ "A", 1/20/2025, "A", 1/21/2025 }`.

**`table`**. For each row of the three columns, use `date_sequence` to generate an array of dates, then `prepend_id` the corresponding id value to each date. Applied to `A2:C4` in the example, this results in a 2D array, in which some items are blank:

    ```
    {{ "A", 1/20/2025, "A", 1/21/2025,    ,          ,    ,          ,    ,          , };
     { "B", 8/8/2025 , "B", 8/9/2025 , "B", 8/10/2025,    ,          ,    ,          , };
     { "C", 12/1/2025, "C", 12/2/2025, "C", 12/3/2025, "C", 12/4/2025, "C", 12/5/2025, }}
    ```

**return**. Use `TOROW` to flatten into a 1D array, ignoring blanks ...

    ```
    { "A", 1/20/2025, "A", 1/21/2025, "B", 8/8/2025 , "B", 8/9/2025 , "B", 8/10/2025, ... }
    ```

... and then use `WRAPROWS` to rearrange into the desired two-column form.

    ```
    {{ "A", 1/20/2025 };
     { "A", 1/21/2025 };
     { "B", 8/8/2025  };
     { "B", 8/9/2025  };
     { "B", 8/10/2025 };
     ... }
    ```

### Example data

Here is a markdown representation of the input.

|       |  *A*  |    *B*     |    *C*     |
| -----:| -----:| ----------:| ----------:|
|  *1*  |   ID  | Date Start | Date End   |
|  *2*  |   A   | 1/20/2025  | 01/21/2025 |
|  *3*  |   B   |  8/8/2025  | 08/10/2025 |
|  *4*  |   C   | 12/1/2025  | 12/05/2025 |
