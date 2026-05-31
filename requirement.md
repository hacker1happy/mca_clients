* Read the create sql queries from the SQL create_mca_client_tables.sql

* Develop a Python script that takes a root folder path as input.

* The root folder will contain a folder Files_to_be_moved_in_DB - where all company folders whose names already exist (might be) in the `COMPANIES` table in column "FOLDER_NAME".

* For every folder-files:
  * Create an entry in the `FILES` table if it does not already exist
  * Insert the file name, with sequential file_id and respective comp_id fetched from the companies table based on the folder_name (it means that the file belonging to the particular folder name is already present in the folder_name col of companies table) & created time
  * for the company not present move the parent company folder to another folder - company_not_found, telling that this companies folders are not present and this companies files can not be moved for now.

* A file should only be processed further if:
  * The file name contains `List of share holders debenture holders`
  * The file extension is `.xlsm`
  * The Excel file contains a sheet named `Sheet-With Validations` and `Sheet-Without Validations`

* If any validation fails:
  * Skip the file
  * Store the file path and proper error message in a log file
  * Move the folder-file to a new unprocessed folder

* If all validations pass:
  * Read data from the either `Sheet-With Validations` or `Sheet-Without Validations` sheet, whichever is non-empty. (We have been also asked to check if at a given point of time only a single sheet should have values else it should be consolidated into one sheet and then move the data). If both the sheets are non-empty then move the data from both the sheets.
  * Extract the following columns:
    * Name of shareholder/ debenture holder (col-E)
    * Folio Number / Reference Number (col-H)
    * DP ID-Client Id-Account Number (col-I)
    * Type of Identifier (col-L)
    * Identification No. (col-M)
    * Occupation (col-N)
    * Number of security held (col-O)
    * Total amount of securities held (in INR) (col-Q)

* Insert the extracted data into the `MCA_CLIENTS` table along with the related `FILE_ID`.

* Commit the transaction after successful insertion of each file’s data.

* Ensure proper exception handling, rollback support, duplicate file checking, and logging throughout the process.

Important notes:
1. assume the immediate child folder under ROOT_FOLDER maps to COMPANIES.FOLDER_NAME 
2. If the same file name exists in different subfolders of the same company, then Should these be treated as: Different files (store relative path instead) 
3. In MCA files, are the column headers always on the first row? -> Yes 
4. Suppose a file was partially loaded previously. then DELETE FROM MCA_CLIENTS WHERE FILE_ID = :file_id and reload.