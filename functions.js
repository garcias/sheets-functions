/**
 * TRIM_HEADER
 *   Remove the first N rows from top of an array
 *   Useful to define an ARRAYFORMULA that includes the header row
 *   but don't want to include it in argument to other functions
 * 
 * @param {Range} full_array - array to trim header from, e.g. A:D
 * @param {Integer} number_of_rows - rows to remove from top, e.g. 1
 * @return {Range} same as full_array with one less row
 */
TRIM_HEADER
= QUERY( 
  full_array, 
  TEXTJOIN( "", FALSE, "select * offset ", number_of_rows ), 
  0 
)

/**
 * REPLACE_HEADER
 *   Replaces the first row of an array with specified headers
 *   Useful when you want to put ARRAYFORMULA calculation into the first row, but
 *   give the column an arbitrary header
 * 
 * @param {Range} full_array - array or generating arrayformula, e.g. A:D
 * @param {Range} headers_array - an array to specify header of each column of full_array; 
 *                                can be a string for a single-column full_array, 
 *                                e.g. A1:B1, {"last", "first"}, "name"
 * @return {Range} same as full_array except for the top row
 */
REPLACE_HEADER
= {
  headers_array; 
  TRIM_HEADER( full_array, 1)
}

/**
 * MELT
 *   Unpivots columns and stacks them. First row of each range assumed to be headers; 
 *   headers of values_range will define variable labels of values in the unpivoted columns.
 *   This implementation is 4-6X faster than MELT_LEGACY
 * 
 * @param {Array} index_range - columns containing index values, including headers; each value in a row
 *                              will be duplicated for each variable in the row, e.g. A:B
 * @param {Array} values_range - 2-D range of values under the variables_header, including headers e.g. C:N
 * @return {Array} an Array of results, see example
 * 
 * @example
 * { "first", "last",  "arrive",  "leave"   ;
 *   "Ann",   "Allen", "8:00 am", "4:00 pm" ;
 *   "Ben",   "Brink", "7:45 am", "3:45 pm" }
 * MELT( A:B, C:D )
 * { "first", "last",  "variable",  "value"   ;
 *   "Ann",   "Allen", "arrive", "8:00 am" ;
 *   "Ann",   "Allen", "leave",  "4:00 pm" ;
 *   "Ben",   "Brink", "arrive", "7:45 am" ;
 *   "Ben",   "Brink", "leave",  "3:45 pm" }
 */
MELT
= LET(
  index_rows, TRIM_HEADER( index_range, 1),
  values_rows, TRIM_HEADER( values_range, 1),
  num_index_rows, ROWS( index_rows ),
  num_labels, COLUMNS( values_range ),
  dup_index_rows, BYCOL( index_rows, LAMBDA( col,
    FLATTEN( WRAPCOLS( FLATTEN( 
      MAKEARRAY( num_labels, 1, LAMBDA( x, y, WRAPROWS( FLATTEN(col), num_index_rows ) ) )
    ), num_index_rows ) )
  ) ),
  labels, FLATTEN( MAKEARRAY( num_index_rows, 1, 
    LAMBDA( x, y, CHOOSEROWS( values_range, 1) ) ) 
  ),
  values, FLATTEN( values_rows ),
  { 
    CHOOSEROWS(index_range, 1), "variable", "value" ; 
    dup_index_rows, labels, values
  }
)

/**
 * MELT_LEGACY
 *   Unpivots columns and stacks them. First row of each range assumed to be headers; 
 *   headers of values_range will define variable labels of values in the unpivoted columns.
 * 
 *   The TRANSPOSE-QUERY-based header-concatenation trick was inspired by Prashanth KV at: 
 *   https://infoinspired.com/google-docs/spreadsheet/unpivot-a-dataset-in-google-sheets-reverse-pivot-formula/
 *   
 *   DEPRECATED. Superseded by the more performant MELT, which uses modern Sheets formulas to
 *   implement the same functionality. Kept to remember the hacky but ingenious string construction and
 *   manipulation methods we used to rely on for reshaping data tables.
 * 
 * @param {Array} index_range - columns containing index values, including headers; each value in a row
 *                              will be duplicated for each variable in the row, e.g. A:B
 * @param {Array} values_range - 2-D range of values under the variables_header, including headers e.g. C:N
 * @return {Array} an Array of results, see example
 * 
 * @example
 * { "first", "last",  "arrive",  "leave"   ;
 *   "Ann",   "Allen", "8:00 am", "4:00 pm" ;
 *   "Ben",   "Brink", "7:45 am", "3:45 pm" }
 * MELT( A:B, C:D )
 * { "first", "last",  "variable",  "value"   ;
 *   "Ann",   "Allen", "arrive", "8:00 am" ;
 *   "Ann",   "Allen", "leave",  "4:00 pm" ;
 *   "Ben",   "Brink", "arrive", "7:45 am" ;
 *   "Ben",   "Brink", "leave",  "3:45 pm" }
 */
MELT_LEGACY
= {
  { CHOOSEROWS(index_range, 1), "variable", "value" }; 
  ARRAYFORMULA( 
    SPLIT( 
      FLATTEN( 
        TRANSPOSE( QUERY( TRANSPOSE( TRIM_HEADER(index_range, 1) & DELIM() ), , 10^100 ) ) & 
        CHOOSEROWS(values_range, 1) & DELIM() & 
        TRIM_HEADER(values_range, 1) 
      ), DELIM(), false, false 
    ) 
  )
}

/**
 * VLOOKUPIFERROR
 *   VLOOKUP wrapped with IFERROR 
 * 
 * @param {value} key - value to search for, e.g. 5, "me@email.com", A2, 
 * @param {Range} datarange - range containing data table (key must be in first column), e.g. sheet1!A:D
 * @param {Integer} index - offset of column that contains result values, e.g. 3
 * @param {Boolean} sorted - whether the keys are sorted, e.g. "FALSE"
 * @param {value} default - value to return if error, e.g. 0, ""
 * @return {value} value that matches key
 */
VLOOKUPIFERROR
= IFERROR( 
  VLOOKUP( key, datarange, index, sorted ),
  default
)

/**
 * VLOOKUPIFERRORARRAY
 *   Applies VLOOKUPIFERROR to a range of keys, with "" as default value
 * 
 * @param {Range} keyrange - range of keys, e.g. A:A
 * @param {Range} datarange - range containing data table (key must be in first column), e.g. sheet1!A:D
 * @param {Integer} index - offset of column that contains result values, e.g. 3
 * @param {Boolean} sorted - whether the keys are sorted, e.g. "FALSE"
 * @return {Array} an Array of results, one for each value in the keyrange
 */
VLOOKUPIFERRORARRAY
= ARRAYFORMULA(
  IFERROR( VLOOKUP( keyrange, datarange, index, sorted ), "" )
)

/**
 * VLOOKUPIFERRORARRAYWITHHEADER
 *   Applies VLOOKUPIFERROR to a range of keys, with "" as default value, with header column
 * 
 * @param {Range} keyrange - range of keys, e.g. A:A
 * @param {Range} datarange - range containing data table (key must be in first column), e.g. sheet1!A:D
 * @param {Integer} index - offset of column that contains result values, e.g. 3
 * @param {Boolean} sorted - whether the keys are sorted, e.g. "FALSE"
 * @param {Text} header - text to display as the column header, e.g. "Last name"
 * @param {Integer} headerrow - row number of header, e.g. 1
 * @return {Array} an Array of results, one for each value in the keyrange
 */
VLOOKUPIFERRORARRAYWITHHEADER
= ARRAYFORMULA(
  IF(
    row(keyrange)=headerrow,
    header,
    IFERROR( VLOOKUP( keyrange, datarange, index, sorted ), "" )
  )
)

/**
 * XLOOKUPIFERROR
 *   XLOOKUP wrapped with IFERROR
 * 
 * @param {value} key - value to search for, e.g. 5, "me@email.com", A2, 
 * @param {Range} lookup_range - range to search for the key, e.g. sheet1!D:D
 * @param {Range} result_range - range that contains result values, e.g. sheet1!A:A
 * @param {value} default - value to return if error, e.g. 0, ""
 * @return {value} value in result_range that matches key
 */
XLOOKUPIFERROR
= IFERROR( 
  XLOOKUP( key, lookup_range, result_range ),
  default
)

/**
 * XLOOKUPIFERRORARRAY
 *   Applies XLOOKUPIFERROR to a range of keys, with "" as default value
 * 
 * @param {Range} keyrange - range of keys, e.g. A:A
 * @param {Range} lookup_range - range to search for each key, e.g. sheet1!D:D
 * @param {Range} result_range - range that contains result values, e.g. sheet1!A:A
 * @return {Array} an Array of results, one for each value in the keyrange
 */
XLOOKUPIFERRORARRAY
= ARRAYFORMULA(
  IFERROR( XLOOKUP( keyrange, lookup_range, result_range ), "" )
)

/**
 * XLOOKUPIFERRORARRAYWITHHEADER
 *   Applies XLOOKUPIFERROR to a range of keys, with "" as default value, with header column
 * 
 * @param {Range} keyrange - range of keys, e.g. A:A
 * @param {Range} lookup_range - range to search for each key, e.g. sheet1!D:D
 * @param {Range} result_range - range that contains result values, e.g. sheet1!A:A
 * @param {Text} header - text to display as the column header, e.g. "Last name"
 * @param {Integer} headerrow - row number of header, e.g. 1
 * @return {Array} an Array of results, one for each value in the keyrange
 */
XLOOKUPIFERRORARRAYWITHHEADER
= ARRAYFORMULA(
  IF(
    row(keyrange)=headerrow,
    header,
    IFERROR( XLOOKUP( keyrange, lookup_range, result_range ), "" )
  )
)

/**
 * CROSSJOIN
 *   Cross joins two arrays as two columns, generating the Cartesian product (every pairwise combination 
 *   of values between two sets). 
 *   WARNING: sheet will start to lag if the result of this function exceeds 100,000 rows; even if as an
 *   intermediate array that gets filtered or array-constrained before the final output
 * 
 * @param {Array} entities_range - array of entities to repeat, e.g. entities!A:A
 * @param {Array} attributes_range - array of attributes per entity, e.g. attributes!A:B
 * @return {Array} an Array of results, one for each value in the keyrange
 */
CROSSJOIN
= LET(
  number_of_entities, rows( entities_range ),
  number_of_attributes, rows( attributes_range ),
  {
    WRAPCOLS(
      FLATTEN( 
        MAKEARRAY( 
          1, 
          number_of_attributes, 
          LAMBDA( row, col, FLATTEN(TRANSPOSE(entities_range)) ) ) 
      ),
      number_of_attributes * number_of_entities
    ),
    WRAPROWS(
      FLATTEN(
        MAKEARRAY( 
          number_of_entities, 
          1, 
          LAMBDA( row, col, TRANSPOSE( FLATTEN(attributes_range)) ) )
      ),
      COLUMNS( attributes_range )
    )
  }
)

/**
 * INNERJOIN
 *   Joins two tables on specified key columns; keys do not have to be unique in each column. First row 
 *   of each range assumed to be headers. "Left" and "right" are arbitrary, but performance is better 
 *   if left table is smaller than right table.
 * 
 *   WARNING: sheet will start to lag if the product of table sizes exceeds 20,000,000, even if 
 *   the final output is only a few rows due to QUERY, FILTER, or ARRAY_CONSTRAIN.
 * 
 * @param {Array} data_left_lbl  - columns of left table to join, may include key column, e.g. orders!A:C
 * @param {Array} keys_left_lbl  - column of left table containing the keys to join on, e.g. orders!B:B
 * @param {Array} data_right_lbl - columns of right table to join, may include key column, e.g. recipes!B:C
 * @param {Array} keys_right_lbl - column of right table containing the keys to join on, e.g. recipes!A:A
 * @return {Array} Array of results such that keys_left == keys_right, see example
 * 
 * @example
  * on sheet named orders
 * { { "diner",    "dish",     "number"  } ;
 *   { "Sengupta", "omelette",    2      } ;
 *   { "Weisz",    "omelette",    1      } ;
 *   { "Weisz",    "pancake",     1      } ;
 *   { "Lemmy",    "waffles",     4      } }
 * 
 * on sheet named recipes
 * { { "dish",     "ingredient",  "amount" } ;
 *   { "omelette", "egg",         "120 g"  } ;
 *   { "pancake",  "egg",         "60 g"   } ;
 *   { "pancake",  "milk",        "40 g"   } ;
 *   { "pancake",  "flour",       "150 g"  } }
 * 
 * = INNERJOIN( orders!A:C, orders!B:B, recipes!B:C, recipes!A:A )
 * { { "diner",    "dish",     "number", "ingredient",  "amount"  } ;
 *   { "Sengupta", "omelette",    2    , "egg"       ,  "120 g"   } ;
 *   { "Weisz",    "omelette",    1    , "egg"       ,  "120 g"   } ;
 *   { "Weisz",    "pancake",     1    , "egg"       ,  "60 g"    } ;
 *   { "Weisz",    "pancake",     1    , "milk"      ,  "40 g"    } ;
 *   { "Weisz",    "pancake",     1    , "flour"     ,  "150 g"   } }
 */
INNERJOIN
= LET(
  data_left, TRIM_HEADER(data_left_lbl,1), data_right, TRIM_HEADER(data_right_lbl,1),
  keys_left, TRIM_HEADER(keys_left_lbl,1), keys_right, TRIM_HEADER(keys_right_lbl,1),
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
  { CHOOSEROWS( data_left_lbl, 1 ), CHOOSEROWS( data_right_lbl, 1); notblank }
)

/**
 * LEFTJOIN
 *   Joins two tables on specified key columns; keys do not have to be unique in each column. First row 
 *   of each range assumed to be headers. "Left" and "right" are NOT arbitrary; rows in the left table 
 *   with no matches will still appear in the resulting array, with blank cells to the right. In this context,
 *   "blank" means IF(TRUE,) rather than "".
 *   
 *   WARNING: sheet will start to lag if the product of table sizes exceeds 20,000,000, even if 
 *   the final output is only a few rows due to QUERY, FILTER, or ARRAY_CONSTRAIN.
 * 
 * @param {Array} data_left_lbl  - columns of left table to join, may include key column, e.g. orders!A:C
 * @param {Array} keys_left_lbl  - column of left table containing the keys to join on, e.g. orders!B:B
 * @param {Array} data_right_lbl - columns of right table to join, may include key column, e.g. recipes!B:C
 * @param {Array} keys_right_lbl - column of right table containing the keys to join on, e.g. recipes!A:A
 * @return {Array} Array of results such that keys_left == keys_right, see example
 * 
 * @example
 * on sheet named orders
 * { { "diner",    "dish",     "number"  } ;
 *   { "Sengupta", "omelette",    2      } ;
 *   { "Weisz",    "omelette",    1      } ;
 *   { "Weisz",    "pancake",     1      } ;
 *   { "Lemmy",    "waffles",     4      } }
 * 
 * on sheet named recipes
 * { { "dish",     "ingredient",  "amount" } ;
 *   { "omelette", "egg",         "120 g"  } ;
 *   { "pancake",  "egg",         "60 g"   } ;
 *   { "pancake",  "milk",        "40 g"   } ;
 *   { "pancake",  "flour",       "150 g"  } }
 * 
 * = LEFTJOIN( orders!A:C, orders!B:B, recipes!B:C, recipes!A:A )
 * { { "diner",    "dish",     "number", "ingredient",  "amount"  } ;
 *   { "Sengupta", "omelette",    2    , "egg"       ,  "120 g"   } ;
 *   { "Weisz",    "omelette",    1    , "egg"       ,  "120 g"   } ;
 *   { "Weisz",    "pancake",     1    , "egg"       ,  "60 g"    } ;
 *   { "Weisz",    "pancake",     1    , "milk"      ,  "40 g"    } ;
 *   { "Weisz",    "pancake",     1    , "flour"     ,  "150 g"   } ;
 *   { "Lemmy",    "waffles",     4    , IF(TRUE,)   ,  IF(TRUE,) } }
 */
LEFTJOIN
= LET(
  data_left, TRIM_HEADER(data_left_lbl,1), data_right, TRIM_HEADER(data_right_lbl,1),
  keys_left, TRIM_HEADER(keys_left_lbl,1), keys_right, TRIM_HEADER(keys_right_lbl,1),
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
  { CHOOSEROWS( data_left_lbl, 1 ), CHOOSEROWS( data_right_lbl, 1) ; notblank }
)

/**
 * QBN
 *   QUERY function but query string refers to columns by names enclosed in backticks ``
 * 
 *   Developed by Stack Exchange user carecki (https://webapps.stackexchange.com/users/304122/carecki)
 *   https://webapps.stackexchange.com/questions/57540/can-i-use-column-headers-in-a-query/167714#167714
 * 
 * @param {Array} data - array whose first row is column headers, e.g. A:F
 * @param {String} query_text - query that refers to column headers in backticks
 *                              e.g. "SELECT `last`, `first` WHERE `reported age` > 30"
 * @return {Array}
 */
QBN
= QUERY(
  {data}, 
  LAMBDA( text, columns,
    REDUCE( 
      text, 
      FILTER( columns, NOT(ISBLANK(columns)) ), 
      LAMBDA( res, col, 
        REGEXREPLACE(res, "`" & col & "`", "Col" & MATCH(col, columns, 0))
      )
    )
  ) ( query_text, ARRAY_CONSTRAIN( data, 1, COLUMNS(data) ) ),
  1
)

/**
 * TESTING 
 *   Formulas useful for generating test data
 */

/**
 * TESTALPHAARRAY
 *   generate array of random strings of lower-case alphabetic characters
 * 
 * @param {Integer} nrows    - height of array to generate
 * @param {Integer} ncolumns - width of array to generate
 * @param {Integer} length  - length of string
 * @return {Array}
 */
TESTALPHAARRAY
= MAKEARRAY( nrows, ncolumns, LAMBDA( x, y,
  JOIN( "", MAKEARRAY( length, 1, LAMBDA( x, y, CHAR(RANDBETWEEN( 97, 122 )) ) ) )
) )

/**
 * TESTASCIIARRAY
 *   generate array of random strings of characters with codes 40-126 in ASCII table
 * 
 * @param {Integer} nrows    - height of array to generate
 * @param {Integer} ncolumns - width of array to generate
 * @param {Integer} length  - length of string
 * @return {Array}
 */
TESTASCIIARRAY
= MAKEARRAY( nrows, ncolumns, LAMBDA( x, y,
  JOIN( "", MAKEARRAY( length, 1, LAMBDA( x, y, CHAR(RANDBETWEEN( 40, 126 )) ) ) )
) )

/**
 * TESTSELECTARRAY
 *   generate array of values selected randomly from an array of options (with replacement)
 * 
 * @param {Integer} nrows    - height of array to generate
 * @param {Integer} ncolumns - width of array to generate
 * @param {Array} options_array - array of values to select from, e.g. A1:B10
 * @return {Array}
 */
TESTSELECTARRAY
= MAKEARRAY( nrows, ncolumns, LAMBDA( x, y,
  LET(
    selection_array, FLATTEN( options_array ),
    number_of_options, ROWS( selection_array ),
    idx, RANDBETWEEN( 1, number_of_options ),
    INDEX( selection_array, idx )
  )
) )

/**
 * TESTSELECTMULTIPLEARRAY
 *   generate array of multiple unique values selected randomly from an array of options
 * 
 * @param {Integer} nrows    - height of array to generate
 * @param {Integer} ncolumns - width of array to generate
 * @param {Array} options_array - array of values to select from, e.g. A1:B10
 * @param {Integer} number_of_selections - max number to select from array, e.g. 4
 * @return {Array} - each value is a string formed as comma separated list of selected options
 */
TESTSELECTMULTIPLEARRAY
= MAKEARRAY( nrows, ncolumns, LAMBDA( x, y,
  LET(
    selections_array, MAKEARRAY( 
      number_of_selections, 1, LAMBDA( x, y, TESTSELECTARRAY( 1, 1, options_array ) ) 
    ),
    JOIN( ", ", unique( selections_array ) )
  )
) )

= MAKEARRAY( nrows, ncolumns, LAMBDA( x, y,
  LET(
    selections_array, TESTSELECTARRAY( number_of_selections, 1, options_array ),
    JOIN( ", ", unique( selections_array ) )
  )
) )
