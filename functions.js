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
 *   Unpivots columns and stacks them. First row of each range assumed to be headers, 
 *   which will define variable labels of values in the unpivoted columns.
 * 
 *   The TRANSPOSE-QUERY-based header-concatenation trick was inspired by Prashanth KV at: 
 *   https://infoinspired.com/google-docs/spreadsheet/unpivot-a-dataset-in-google-sheets-reverse-pivot-formula/
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
 *   Generates the Cartesian product of two arrays as a two-column array of pairs
 *   A cross join (Cartesion product) is every pairwise combination of values between two sets.
 * 
 * @param {Array} entities_range - 1-D array of entities to repeat, e.g. entities!A:A
 * @param {Array} attributes_range - 1-D array of attributes per entity, e.g. attributes!A:A
 * @return {Array} an Array of results, one for each value in the keyrange
 */
CROSSJOIN
= {
  TRANSPOSE( 
    SPLIT( 
      JOIN( DELIM(), 
        ARRAYFORMULA(
          REPT( entities_range & DELIM(), ROWS(attributes_range) )
        )
      ), DELIM()
    )
  ),
  TRANSPOSE( 
    SPLIT(
      REPT(
        JOIN( DELIM(), attributes_range ) & DELIM(), ROWS(entities_range) 
      ), DELIM()) 
  )
}

/**
 * QBN
 *   QUERY function but query string refers to columns by names enclosed in backticks ``
 * 
 *   Developed entirely by Stack Exchange user carecki (https://webapps.stackexchange.com/users/304122/carecki)
 *   https://webapps.stackexchange.com/questions/57540/can-i-use-column-headers-in-a-query/167714#167714
 * 
 * @param {Array} data - array whose first row is column headers, e.g. A:F
 * @param {String} query_text - query that refers to column headers in backticks
 *                              e.g. "select `last`, `first` where `reported age` > 30"
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
