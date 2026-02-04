/**
 * TRIM_HEADER_LEGACY
 *   Remove the first N rows from top of an array
 *   Useful to define an ARRAYFORMULA that includes the header row
 *   but don't want to include it in argument to other functions
 * 
 * @param {Range} full_array - array to trim header from, e.g. A:D
 * @param {Integer} number_of_rows - rows to remove from top, e.g. 1
 * @return {Range} same as full_array with one less row
 */
TRIM_HEADER_LEGACY
= QUERY( 
  full_array, 
  TEXTJOIN( "", FALSE, "select * offset ", number_of_rows ), 
  0 
)

/**
 * REPLACE_HEADER_LEGACY
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
REPLACE_HEADER_LEGACY
= {
  headers_array; 
  TRIM_HEADER( full_array, 1)
}

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
