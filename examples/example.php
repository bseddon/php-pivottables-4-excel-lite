<?php

use \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Groups;
use \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Group;
use \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Spreadsheet;

require_once __DIR__ . '/data.php';
require_once __DIR__ . '/../vendor/autoload.php';

if ( ! class_exists( "\lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Spreadsheet", false ) )
{
	require_once __DIR__ . '/../phpspreadsheet/Spreadsheet.php';
}

$outputFileName = __DIR__ . '/generated.xlsx';

$spreadsheet = new Spreadsheet();

$spreadsheet->getProperties()
	->setCreator("XBRL Query Generator")
	->setLastModifiedBy("XBRL Query Generator")
	->setTitle("Microsoft 2018 QK")
	->setSubject("Pivot table report")
	->setDescription("This could be an explanation")
	->setKeywords("xbrl microsoft 2018 10k")
	->setCategory("Reports");

$data = load_data();

$networks = array(
	// All pivot tables are added to sheets to which the data is added starting at cell B2.

	// The first PT is added to a sheet called 'Worksheet'.  It has two groups on the rows
	// (Account/Genre) that are filtered to three of the accounts.  Because there is filtering
	// the sort type must be 'manual'.  There are no groups added to the columns.  Instead,
	// the columns are the values of three numeric columns.
	// Note that while the row groups object is created passing an explicit 'Group' instance
	// the value groups instance is created by passing an array of string names.  This is a
	// simple ay to create groups if the default group values (no filtering and sort type
	// ascending) are acceptable.
	array(	'data' => $data,
			'args' => array(
				"Worksheet",
				2 + count( $data ) + 1 + 3, 2,
				new Groups( array( new Group( 'Account', 'manual', array( 'Megan', 'Daniel', 'Hannah' ) ), 'Genre' ) ),
				new Groups(),
				new Groups( array( 'Total Size', 'Images', 'Average Ranking' ) )
			)
	),

	// The second PT is added to a sheet called 'Worksheet2'.  It has just one group on the rows
	// (Account) that is not filtered but the account names will be displayed in descending order.
	// This PT has two groups on the columns (Genre/Images).  The values are from the 'Total Size'
	// column.
	array(	'data' => $data,
			'args' => array(
				"Worksheet2",
				2 + count( $data ) + 1 + 3, 2,
				new Groups( new Group( 'Account', 'descending' ) ),
				new Groups( array( 'Genre', 'Images' ) ),
				new Groups( array( 'Total Size' ) )
			)
	),

	// The third PT is added to a sheet called 'Worksheet3'. It has two groups on the rows
	// (Account/Genre).  There is one group on the columns (Images) and the values are from the
	// 'Total Size' column.
	array(	'data' => $data,
			'args' => array(
				"Worksheet3",
				2 + count( $data ) + 1 + 3, 2,
				new Groups( array( 'Account', 'Genre' ) ) ,
				new Groups( array( 'Images' ) ),
				new Groups( array( 'Total Size'  ) )
			)
	)
);

foreach ( $networks as $index => $network )
{
	$range = $spreadsheet->addData( $data, $network['args'][0] );
	$spreadsheet->addNewPivotTable( $data, $range, ...$network['args'] );
}

$writer = PhpOffice\PhpSpreadsheet\IOFactory::createWriter( $spreadsheet, $inputFileType );
$writer->save($outputFileName);

$spreadsheet->disconnectWorksheets();
unset( $spreadsheet );
