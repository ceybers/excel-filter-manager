# excel-filter-manager
Stores and restores the state of table filters in Excel

## Filter Manager

Serializes and de-serializes the common filter types.

The criteria are converted to Base64. The parameters for the filter for each columns are stored in a simple comma separated format. Then all the filters are stored separated by semicolons.

This allows us to copy and paste to clipboard, or to store it in a cell in a hidden worksheet to store/restore the filter, via a VBA form (not yet implemented).

## TODO

* Proper UI for storing and restoring filters
* Quick toggle on/off for filters on active table
* Selectively store/restore filters on a column-by-column basis
* Store filter state in a more sensible format than the current comma and semi-colon format
* Replace current test code with proper unit tests