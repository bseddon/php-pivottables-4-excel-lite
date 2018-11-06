# php-pivottables-4-excel-lite

PHPOffice/PhpSpreadsheet is a great project to read and write Excel workbook but it does not support some Excel features such as Tables and Pivot Tables.  This project extends PhpSpreadsheet by adding support for pivot tables but only in a limited way.

## What is supported?
This project ensures that existing pivot tables are retained and allows pivot tables to be created to report on data in worksheets.  The rows and columns can be defined based on columns in the worksheet and they can be filtered and sorted.

## What is not supported?
The pivot table features not supported include:
- External data sources
- Styling
- Hierarchies
- Formulas
However, there is no reason why support for these features cannot be added and the project shows how additional features can be implemented.

# Installing
Use composer with the command:
composer install lyquidity/php-pivottables-4-excel-lite:dev-master --prefer-dist

# Getting started

The ./examples/example.php file includes illustrations of using the classes.

Assuming you have installed the library using composer then this PHP application will run the test:

```php
<?php
require_once __DIR__ . '/vendor/autoload.php';
$spreadsheet = new Spreadsheet();
$data = array(..data..);
$range = $spreadsheet->addData( $data, 'Workbook Title' );
$spreadsheet->addNewPivotTable( 
   $data, $range, 2, 2, 
   new Groups( array( 'Row1', 'RowN' ) ), 
   new Groups( array( 'Col1', 'ColN' ) ), 
   new Groups( array( 'Value1', 'ValueN' ) )
);
```
