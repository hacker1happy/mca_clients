* We already have a `COMPANIES` table with `SEQ_NO`, write sql query to set it as as the primary key.

* Create a new `FILES` table to store file-related details:

  * `FILE_NUMBER` as Primary Key
  * `FILE_NAME`
  * `COMPANY_NUMBER` as Foreign Key referencing `COMPANIES.SEQ_NO`

* Create a new `CLIENTS` table to store shareholder/debenture-holder data from Excel files:

  * `CLIENT_SEQUENCE_NUMBER` as Primary Key
  * `NAME_OF_SHAREHOLDER`
  * `FOLIO_REFERENCE_NUMBER`
  * `DP_CLIENT_ACCOUNT`
  * `TYPE_OF_IDENTIFIER`
  * `IDENTIFICATION_NO`
  * `OCCUPATION`
  * `NUMBER_OF_SECURITY_HELD`
  * `TOTAL_SECURITY_AMOUNT_INR`
  * `FILE_NUMBER` as Foreign Key referencing `FILES.FILE_NUMBER`
  * `CREATED_AT`

* Develop a Python script that takes a root folder path as input.

* The root folder will contain company folders whose names already exist in the `COMPANIES` table.

* Recursively scan all folders and subfolders.

* If any `.zip` file is found:

  * Extract it into the same parent folder
  * Delete the original `.zip` file after extraction

* For every extracted or existing file:

  * Create an entry in the `FILES` table if it does not already exist

* A file should only be processed further if:

  * The file name contains `List of share holders debenture holders`
  * The file extension is `.xlsm`
  * The Excel file contains a sheet named `Sheet-With Validations`

* If any validation fails:

  * Skip the file
  * Store the file path and proper error message
  * Generate a final JSON error report containing all skipped files and reasons

* If all validations pass:

  * Read data from the `Sheet-With Validations` sheet
  * Extract the following columns:

    * Name of shareholder/ debenture holder
    * Folio Number / Reference Number
    * DP ID-Client Id-Account
    * Type of Identifier
    * Identification No.
    * Occupation
    * Number of security held
    * Total amount of securities held (in INR)

* Insert the extracted data into the `CLIENTS` table along with the related `FILE_NUMBER`.

* Commit the transaction after successful insertion of each file’s data.

* Ensure proper exception handling, rollback support, duplicate file checking, and logging throughout the process.
