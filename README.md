# PHP Pivot Tables for Excel Lite

PHPOffice/PhpSpreadsheet is a great project to read and write Excel workbook but it does not support some Excel features such as 
Tables and Pivot Tables.  This project extends PhpSpreadsheet by adding support for pivot tables but only in a limited way.

## What is supported?
This project ensures that existing pivot tables are retained and allows pivot tables to be created to report on data in worksheets.  
The rows and columns can be defined based on columns in the worksheet and they can be filtered and sorted. Only Xlsx/Xlsm files are
supported.

## What is not supported?
The pivot table features not supported include:
- External data sources
- Styling
- Hierarchies
- Formulas
- File types other than Xlsx

However, there is no reason why support for these features cannot be added and the project shows how additional features can be implemented.

# Installing
Use composer with the command:

```
composer install lyquidity/php-pivottables-4-excel-lite:dev-master --prefer-dist
```

# Getting started

The ./examples/example.php file includes illustrations of using the classes.

Assuming you have installed the library using composer then this PHP application will run the test:

```php
<?php
require_once __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/vendor/lyquidity/php-pivottables-4-excel-lite/examples/example.php';
```

The examples use the following simple data set:

|Account|Genre|Images|Average Ranking|Total Size|
|:---|:---|---:|---:|---:|
|Megan	  |Portraits	|20	|4	|72000|
|Hannah	  |Landscapes	|31	|3.5|83000|
|Vicky	  |Floral	    |25	|4.2|42000|
|Ian	  |Portraits	|40	|3.7|92000|
|Michael  |Landscapes	|23	|3.8|72000|
|Daniel   |Landscapes	|29	|4.4|85000|

# Overridden Classes

To implement support for pivot tables it has been necessary to override 5 classes:

|Class|Reason|
|:---|:---|
|Spreadsheet|Extends the PhpSpreadsheet class to add functions that carry forward existing pivot tables and add new ones.  Only addData and addNewPivotTable should be called from your code. The class also maintains a list of the cache definitions, record sets and pivot table definitions.|
|XlsxReader|Registered by the replacement spreadsheet class to handle reading Xlsx documents so that existing pivot table resources can be recorded in a spreadsheet class instance.  The whole XlsxReader class is replicated because it relies on private functions that cannot be called from descendant instances.|
|XlsxWriter|Registered by the replacement spreadsheet class to handle writing Xlsx documents so that pivot table resources recorded in a spreadsheet class instance can be included in the generated package file.  The whole XlsxWriter class is replicated because it relies on private functions that cannot be called from descendant instances.|
|Rels|Add support for the relationships required for pivot table support.  WriteRelationship and writeUnparsedRelationship are reimplemented because they are private in the parent Rels class and cannot be called from this claSS.|
|Workbook|This class is replaced so the <PivotCaches> element can be written.  The whole class is reimplemented because all the functions are private so it is not possible to replace just one.|

# New classes

In addition eight new classes are added:

|Class|Reason|
|:---|:---|
|PivotCacheDefinition|Used to represent a cache definition file in the workbook document|
|PivotCacheDefinitionCollection|Represents the list of existing and new cache definition files|
|PivotCacheRecords|Represents one of the cache records files in the workbook document|
|PivotCacheRecordsCollection|Represents the list of existing and new cache records files|
|PivotTable|Used to represent a pivot table definition file in the workbook document|
|PivotTableCollection|Represents the list of existing and new pivot table definition files|
|Group|Represents a specfic column, row or value field from the data set and is used to define the use of the field in the pivot table.  The class defines the field name, the sort order and any filter applied.|
|Groups|Represents a collection of Group instances to build up the fields use for rows, columns and values|
