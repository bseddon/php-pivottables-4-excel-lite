<?php

namespace lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet;

use \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheDefinition;
use \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheDefinitionCollection;
use \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheRecordsCollection;
use \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheRecords;
use \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotTableCollection;
use \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotTable;
use \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use \PhpOffice\PhpSpreadsheet\Shared\XMLWriter;
use \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Writer\ZipArchiveX;
use \PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use \TupleDictionary;

require_once __DIR__ . '/XlsxWriter.php';
require_once __DIR__ . '/XlsxReader.php';

\PhpOffice\PhpSpreadsheet\IOFactory::registerReader( 'Xlsx', Reader\Xlsx::class );
\PhpOffice\PhpSpreadsheet\IOFactory::registerWriter( 'Xlsx', Writer\Xlsx::class );

require_once 'Xlsx/PivotCacheDefinition.php';
require_once 'Xlsx/PivotCacheDefinitionCollection.php';
require_once 'Xlsx/PivotCacheRecords.php';
require_once 'Xlsx/PivotCacheRecordsCollection.php';
require_once 'Xlsx/PivotTable.php';
require_once 'Xlsx/PivotTableCollection.php';

/**
 * A class to hold group values
 * This class implements IteratorAggregate so it can be used
 * as an iterator in a foreach loop to iterate over groups
 */
class Groups implements \IteratorAggregate
{
	/**
	 * A list of the currently added columns
	 * @var array
	 */
	private $groups = array();

	/**
	 * Add one or more values
	 * @param array|string $values
	 */
	public function __construct( $groups = null )
	{
		if ( ! $groups ) return;

		$this->addGroups( $groups );
	}

	/**
	 * Add one or more groups
	 * @param array|Group $groups
	 */
	public function addGroups( $groups )
	{
		if ( is_string( $groups ) )
		{
			$groups = new Group( $groups );
		}

		if ( ! is_array( $groups ) )
		{
			$groups = array( $groups->getName() => $groups );
		}
		else
		{
			$groups = array_reduce( $groups, function( $carry, $group )
			{
				if ( is_string( $group ) )
				{
					$group = new Group( $group );
				}
				$carry[ $group->getName() ] = $group;

				return $carry;
			}, array() );
		}

		$this->groups = array_merge( $this->groups, $groups );
	}

	/**
	 * Get the group for a name
	 * @return Group
	 */
	public function getGroupByName( $name )
	{
		if ( ! isset( $this->groups[ $name ] ) ) return null;
		return $this->groups[ $name ];
	}

	/**
	 * Get the list of groups
	 * @return string[]
	 */
	public function getGroups()
	{
		return $this->groups;
	}

	/**
	 * Get the list of groups
	 * @return string[]
	 */
	public function getGroupNames()
	{
		return array_map( function( /** @var Group $group */ $group )
		{
			return $group->getName();
		}, $this->groups );
	}

	/**
	 * Get the count of groups
	 * @return number
	 */
	public function count()
	{
		return count( $this->groups );
	}

	/**
	 * Implements the IteratorAggregator interface member
	 * {@inheritDoc}
	 * @see IteratorAggregate::getIterator()
	 */
	public function getIterator()
	{
		return (function ()
		{
			reset($this->groups);
			while(list($key, $val) = each($this->groups))
			{
				yield $key => $val;
			}
		})();
	}
}

/**
 * A class to hold groups
 * This class implements IteratorAggregate so it can be used
 * as an iterator in a foreach loop to iterate over values
 */
class Group implements \IteratorAggregate
{
	/**
	 * The name of the group
	 * @var string
	 */
	private $name;

	/**
	 * A sort type ascending|descending|manual
	 * @var string
	 */
	private $sortType = 'ascending';

	/**
	 * A list of the values to be hidden or visible.
	 */
	private $values = array();

	/**
	 * Whether values in the list should be hidden or visible
	 */
	private $visibleValues = true;

	/**
	 * Constructor
	 * @param string $name		The name of the group
	 * @param string $sortType	One of 'ascending', descending' or 'manual'
	 * @param array $value		An explicit list of the values that should be visible (all others will be hidden)
	 * 							Leave empty to show all values
	 */
	public function __construct( $name, $sortType = 'ascending', $values = array() )
	{
		if ( ! is_string( $name ) ) throw new \Exception("The group name argument is not a string");

		$this->name = $name;
		$this->sortType = $sortType;
		$this->addVisibleValues( $values );
	}

	/**
	 * Return the name of the constructor
	 * @return string
	 */
	public function getName()
	{
		return $this->name;
	}

	/**
	 * A list of the values to be hidden.  All others will be visible.
	 * Any explicitly visible values will be removed.
	 * @param array|string $values
	 * @return Group
	 */
	public function addHiddenValues( $values )
	{
		$this->addVisibleValues( $values );
		$this->visibleValues = ! $values; // Always true if there no values

		return $this;
	}

	/**
	 * A list of the values to be visible.  All others will be hidden.
	 * Any explicitly hidden values will be removed.
	 * Explicitly set values will mean that the sort type will be set to 'manual'
	 * @param array|string $values
	 * @return Group
	 */
	public function addVisibleValues( $values )
	{
		$this->visibleValues = true;
		if ( ! $values )
		{
			$values = array();
		}
		else if ( ! is_array( $values ) )
		{
			$values = array( $values );
		}

		$this->values = $values;
		if ( $values )
		{
			$this->sortType = 'manual';
		}

		return $this;
	}

	/**
	 * Return the visible state of a specific value
	 * @param unknown $value
	 * @return boolean
	 */
	public function getVisibleState( $value )
	{
		if ( ! $value || ! $this->values ) return true;

		return in_array( $value, $this->values )
			? $this->visibleValues
			: ! $this->visibleValues;
	}

	/**
	 * Set the sort type to one of 'ascending', 'descending' or 'manual'
	 * @param string $type
	 */
	public function setSortType( $type )
	{
		if ( ! in_array( $type, array( 'ascending', 'descending' ,'manual' ) ) )
			throw \Exception( "The sort type MUST be one of 'ascending', 'descending' or 'manual'" );

		$this->sortType = $type;
	}

	/**
	 * Get the sort type
	 * @return string
	 */
	public function getSortType()
	{
		return $this->sortType;
	}

	/**
	 * Get the count of groups
	 * @return number
	 */
	public function count()
	{
		return count( $this->values );
	}

	/**
	 * Implements the IteratorAggregator interface member
	 * {@inheritDoc}
	 * @see IteratorAggregate::getIterator()
	 */
	public function getIterator()
	{
		return (function ()
		{
			reset($this->values);
			while(list($key, $val) = each($this->values))
			{
				yield $key => $val;
			}
		})();
	}
}

/**
 * Implements a replacement spreadheet that supports pivot tables
 */
class Spreadsheet extends \PhpOffice\PhpSpreadsheet\Spreadsheet
{
	/**
	 * Path/Xml for references in workbook.xml.rels
	 * @var PivotCacheDefinitionCollection
	 */
	public $pivotCacheDefinitionCollection;

	/**
	 * Path/Xml for references in workbook.xml.rels
	 * @var PivotCacheRecordsCollection
	 */
	public $pivotCacheRecordsCollection;

	/**
	 * Path/Xml for references in workbook.xml
	 * @var PivotTableCollection
	 */
	public $pivotTables;

	// Default constructor
	public function __construct()
	{
		parent::__construct();
		$this->pivotCacheDefinitionCollection = new PivotCacheDefinitionCollection();
		$this->pivotCacheRecordsCollection = new PivotCacheRecordsCollection();
		$this->pivotTables = new PivotTableCollection();
	}

	/**
	 * A function to create and add an existing cache definition
	 * @param string $rId The existing reference Id
	 * @param string $path A normalized path in the Zip file to the existing Xml
	 * @param string $xml The existing Xml
	 * @return \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheDefinition
	 */
	public function addPivotCacheDefinition( $rId, $path, $xml )
	{
		$definition = new PivotCacheDefinition( $path, $xml, $rId );
		$this->pivotCacheDefinitionCollection->addPivotCacheDefinition( $definition );
		return $definition;
	}

	/**
	 * A function to create and add an existing pivot table record set instance
	 * @param string $rId  The existing reference Id
	 * @param string $path A normalized path in the Zip file to the existing Xml
	 * @param string $xml The existing Xml
	 * @param string $cacheDefinitionPath The path to the cache definition that points to this record set
	 * @return \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheRecords
	 */
	public function addPivotCacheRecords( $rId, $path, $xml, $cacheDefinitionPath = null )
	{
		$records = new PivotCacheRecords( $path, $xml, $rId, $cacheDefinitionPath );
		$this->pivotCacheRecordsCollection->addPivotCacheRecords( $records );
		$definition = $this->pivotCacheDefinitionCollection->getPivotCacheDefinitionByPath( $cacheDefinitionPath );
		if ( $definition )
		{
			$definition->addCacheRecords( new PivotCacheRecordsCollection( $records ) );
		}
		return $records;
	}

	/**
	 * A function to create and add an existing pivot table instance
	 * @param string $rId  The existing reference Id
	 * @param string $path The path in the Zip file to the existing Xml
	 * @param string $xml The existing Xml
	 * @param string $cacheId The cache id used in workbook.xml
	 * @param string $sheetName The name of the sheet that refers to this set of records
	 * @return \lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotTable
	 */
	public function addPivotTable( $rId, $path, $xml, $cacheId, $sheetName = null )
	{
		$table = new PivotTable( $path, $xml, $rId, $cacheId, $sheetName );
		$this->pivotTables->addPivotTable( $table );
		return $table;
	}

	/**
	 * Apply the data in the array to $sheetIndex.  If the sheet is missing, one will be added.
	 * @param array $data An associative array of data to add to the sheet
	 * @param int $sheetIndex Can be in number of the sheet or a name
	 * @param number $rowIndex (optional: default=2) The top of the array to populate
	 * @param number $colIndex (optional: default=2) The left of the array to populate
	 * @param bool $hideRows (optional: default=true) When true the rows used by the data will be hidden
	 * @return string The range containing the added data
	 */
	public function addData( $data, $sheetIndex, $rowIndex = 2, $colIndex = 2, $hideRows = true )
	{
		$sheet = $this->getSheetFromIndex( $sheetIndex );

		// $data is an array in which row zero is a list of headers
		$headers = array_keys( $data[0] );
		$values = array_merge( array( $headers ), array_values( $data ) );

		$sheet->fromArray( $values, null, \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex( $colIndex ) . $rowIndex );
		$sheet->getStyleByColumnAndRow( $colIndex, $rowIndex, $colIndex + count( $headers ) - 1, $rowIndex )->getFont()->setBold( true );
		$sheet->setSelectedCell("A1");

		if ( $hideRows )
		foreach ( range( $rowIndex, $rowIndex + count( $values ) - 1 ) as $iRow )
		{
			$sheet->getRowDimension( $iRow )->setVisible( false );
		}

		return	\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex( $colIndex ) . $rowIndex . ":" .
				\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex( $colIndex + count( $headers ) - 1 ) .
				( $colIndex + count( $values ) - 1 );
	}

	private $definitionId = 0;

	/**
	 * Apply the data in the array to $sheetIndex.  If the sheet is missing, one will be added.
	 * A check will be made for other pivot table and an error will be raised if there is an overlap.
	 * @param array $data The data to pivot
	 * @param string $dataRange A string defining a range containing the data from which to create a pivot table
	 * @param int $sheetIndex Can be in number of the sheet or a name
	 * @param number $rowIndex (optional: default=2) The top of the array to populate
	 * @param number $colIndex (optional: default=2) The left of the array to populate
	 * @param Groups $rowGroups The names of fields that should be shown as columns
	 * @param Groups $columnGroups The names of fields that should be shown as columns
	 * @param Groups $valueGroups The names of fields that should be shown as columns
	 * @param string $name The name to use for the pivot table
	 * @return bool True if the pivot table has been created successfully
	 */
	public function addnewPivotTable( $data, $dataRange, $sheetIndex, $rowIndex = 2, $colIndex = 2, $rowGroups = null, $columnGroups = null, $valueGroups = null, $name = "PivotTable1", $dataCaption = 'Data' )
	{
		// Check for existing pivot tables

		// Create a cache definition

		$sheet = $this->getSheetFromIndex( $sheetIndex );
		$date = \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel( time() );
		// $definitionUniqueId = uniqid();
		$this->definitionId++;
		$definitionUniqueId = $this->definitionId;
		$sharedItems = $this->createCacheDefinition( $sheet, $dataRange, $definitionUniqueId, $date );

		// Create a record set
		if ( ! $this->createRecordSet( $sheet, $dataRange, $definitionUniqueId, $sharedItems ) ) return false;

		// Get the count of the data rows
		$rowCount = count( $data );

		$cacheId = PivotCacheDefinition::getNextCacheId();
		$this->pivotCacheDefinitionCollection->addPivotCacheIndex( $cacheId, "rId$definitionUniqueId" );

		// Create the pivot table.
		// The table is place '7' rows after the end of the data to allow for the paging fields.
		if ( ! $this->createPivotTable(
			$sheet, $dataRange, $definitionUniqueId,
			$cacheId, $sharedItems, $sheetIndex, $colIndex, $rowIndex,
			$rowGroups, $columnGroups, $valueGroups, $name, $dataCaption )
		) return false;

		return true;
	}

	/**
	 * Return a sheet that previously exists or a new one if the index does not already exist
	 * @param int|string $sheetIndex
	 * @return NULL|\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
	 */
	private function getSheetFromIndex( $sheetIndex )
	{
		$sheet = null;
		if ( is_numeric( $sheetIndex ) )
		{
			// If the sheet exists, return it otherwise create it
			try
			{
				$sheet = $this->getSheet( $sheetIndex );
			}
			catch( \Exception $ex )
			{
				$sheet = $this->createSheet( $sheetIndex );
			}
		}
		else
		{
				// If the sheet exists, return it otherwise create it
			if ( $this->getSheetByName( $sheetIndex ) )
			{
				$sheet = $this->getSheetByName( $sheetIndex );
			}
			else
			{
				$sheet = $this->createSheet();
				$sheet->setTitle( $sheetIndex, false );
			}

		}

		return $sheet;
	}

	/**
	 * Generate the Xml for a cache definition and return a set of unique values for each column
	 * It is assumed the headers are in the first row
	 * @param Worksheet $sheet
	 * @param string $dataRange
	 * @param string $uniqueId
	 * @param float $date
	 * @return array Unique values for each column
	 */
	private function createCacheDefinition( $sheet, $dataRange, $uniqueId, $date )
	{
		/*
			For more information about the pivotCacheDefinition element see section 18.10.1.67 of
			Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf

			<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<pivotCacheDefinition
			    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
			    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
				r:id="rId1"
				refreshedBy="XBRL Query"
				refreshedDate="41934.8125"
				createdVersion="1"
				refreshedVersion="1"
				recordCount="6"
				refreshOnLoad="1"
			>
			    <cacheSource type="worksheet">
			        <worksheetSource ref="B2:F8" sheet="Data 1"/>
			    </cacheSource>
			    <cacheFields count="5">
			        <cacheField name="Account" numFmtId="0">
			            <sharedItems count="6">
			                <s v="Megan"/>
							...
			            </sharedItems>
			        </cacheField>
			        <cacheField name="Images" numFmtId="0">
			            <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="20" maxValue="40" count="6">
			                <n v="20"/>
							...
			            </sharedItems>
			        </cacheField>
					...
			    </cacheFields>
			</pivotCacheDefinition>
		 */

		$boundary = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries( $dataRange );
		$leftCol = $boundary[0][0];
		$topRow = $boundary[0][1];

		// Access the data
		$data = $sheet->rangeToArray( $dataRange, null, false );
		$headers = array_shift( $data );
		$sharedItems = array();

        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
		$objWriter->startElement('pivotCacheDefinition');
			$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
			$objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
			$objWriter->writeAttribute('r:id', "rId$uniqueId" );
			$objWriter->writeAttribute('refreshedBy', $sheet->getParent()->getProperties()->getCreator());
			$objWriter->writeAttribute('refreshedDate', $date);
			$objWriter->writeAttribute('createdVersion', '1');
			$objWriter->writeAttribute('refreshedVersion', '1');
			$objWriter->writeAttribute('refreshOnLoad', '1');
			$objWriter->writeAttribute('recordCount', count( $data ) );

			$objWriter->startElement('cacheSource');
			$objWriter->writeAttribute('type', 'worksheet' );

				$objWriter->startElement('worksheetSource');
					$objWriter->writeAttribute('ref', $dataRange );
					$objWriter->writeAttribute('sheet', $sheet->getTitle() );

	        	$objWriter->endElement();

			$objWriter->endElement();

			$objWriter->startElement('cacheFields');
				$objWriter->writeAttribute('count', count( $headers ) );

				foreach ( $headers as $index => $header )
				{
					$objWriter->startElement('cacheField');
						$objWriter->writeAttribute('name', $header );
						$objWriter->writeAttribute('numFmtId', 0 );

						// Accumulate a list of unique items for this header
						$type = null;
						$columnValues = array();

						foreach ( $data as $row => $values )
						{
							$columnValues[] = $values[ $index ];

							$cellType = $sheet->getCellByColumnAndRow( $leftCol + $index, $topRow + $row + 1 )->getDataType();
							if ( $type )
							{
								if ( $cellType != $type ) $cellType = "s";
							}
							else $type = $cellType;
						}

						$columnValues = array_values( array_unique( $columnValues ) );

						$objWriter->startElement('sharedItems');
							if ( $type != "s" )
							{
								$nonIntegers = array_filter( $columnValues, function( $value )
								{
									return intval( $value ) != $value;
								} );

								// TODO This needs to be improved.  Hardcoding no mixed types and no always integers and floats is not right.
								$objWriter->writeAttribute('containsSemiMixedTypes', 0 );
								$objWriter->writeAttribute('containsString', 0 );
								$objWriter->writeAttribute('containsNumber', 1);
								$objWriter->writeAttribute('containsInteger', $nonIntegers ? 0 : 1 ); // Truthy value means 'all numbers are integers'
								$objWriter->writeAttribute('minValue', min( $columnValues ) );
								$objWriter->writeAttribute('maxValue', max( $columnValues ) );
							}
							$objWriter->writeAttribute('count', count( $columnValues ) );

							foreach ( $columnValues as $row => $value )
							{
								$objWriter->startElement( $type );
									$objWriter->writeAttribute('v', $value );
		        				$objWriter->endElement();
							}

		        		$objWriter->endElement();

						$sharedItems[ $header ] = $columnValues;

	        		$objWriter->endElement();
				}

        	$objWriter->endElement();

		$objWriter->endElement();

		$zip = new ZipArchiveX();
        $xml = $zip->formattedXml( $objWriter->getData() );
		// echo $xml;
		$this->addPivotCacheDefinition( "rId$uniqueId", "xl/pivotCache/pivotCacheDefinition$uniqueId.xml", $xml );
        return $sharedItems;
	}

	/**
	 * Generate the record set for the cache
	 * It is assumed the headers are in the first row
	 * @param Worksheet $sheet
	 * @param string $dataRange
	 * @param string $definitionUniqueId
	 * @param float $date
	 * @return bool True if the record set is created successfully
	 */
	private function createRecordSet( $sheet, $dataRange, $definitionUniqueId, $sharedItems )
	{
		/**
			<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<pivotCacheRecords
			    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
			    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="6">
			    <r>
			        <x v="0"/>
			        <x v="0"/>
			        <x v="0"/>
			        <x v="0"/>
			        <x v="0"/>
			    </r>
			    ...
			</pivotCacheRecords>
		 */

		$boundary = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries( $dataRange );
		$leftCol = $boundary[0][0];
		$topRow = $boundary[0][1];
		$rightCol = $boundary[1][0];
		$bottomRow = $boundary[1][1];

		$data = $sheet->rangeToArray( $dataRange, null, false );
		$headers = array_shift( $data );

        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // record set
		$objWriter->startElement('pivotCacheRecords');
			$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
			$objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
			$objWriter->writeAttribute('count', $bottomRow - $topRow);

			foreach ( $data as $row => $values )
			{
				$objWriter->startElement('r');

				foreach( $headers as $col => $header )
				{
					$value = $values[ $col ];
					$index = array_search( $value, $sharedItems[ $header ] );

					if ( $index === false )
					{
						// Get the type
						$cellType = $sheet->getCellByColumnAndRow( $leftCol + $col, $topRow + $row + 1 )->getDataType();
						$objWriter->startElement( $cellType );
							$objWriter->writeAttribute('v', $value );
						$objWriter->endElement();
					}
					else
					{
						$objWriter->startElement('x');
							$objWriter->writeAttribute('v', $index );
						$objWriter->endElement();
					}
				}

				$objWriter->endElement();
			}

		$objWriter->endElement();

		$zip = new ZipArchiveX();
		$xml = $zip->formattedXml( $objWriter->getData() );
		// echo $xml;
		$this->addPivotCacheRecords( "rId$definitionUniqueId", "xl/pivotCache/pivotCacheRecords$definitionUniqueId.xml", $xml, "xl/pivotCache/pivotCacheDefinition$definitionUniqueId.xml" );
		return true;
	}

	/**
	 * Return an array of integer offset that correspond to the group position in $columns
	 * @param array $columns A list of columns in shared items order
	 * @param Groups|null $groups A list of groups by name
	 * @return int[]
	 */
	private function createGroupIndexes( $columns, $groups = null )
	{
		return array_keys( array_intersect( $columns, $groups->getGroupNames() ) );

		$groupIndexes = array();

		if ( $groups )
		{
			foreach ( $groups->getGroupNames() as $groupColumn )
			{
				$groupIndexes[] = array_search( $groupColumn, $columns );
			}
		}
		else
		{
			$groupIndexes[] = key( $columns );
		}

		return $groupIndexes;
	}

	/**
	 * Create a list of columns that meet the selection criteria
	 * @param array $headers
	 * @param array $data
	 * @param Groups $rowGroups
	 * @param Groups $sumColumns
	 * @param string $colsAreValues
	 * @return \TupleDictionary
	 */
	private function pivotGroups( $headers, $data, $rowGroups, $sumColumns = null )
	{
		$result = new \TupleDictionary();
		$sumColumnIndexes = array();

		if ( $sumColumns )
		{
			foreach ( $sumColumns->getGroupNames() as $sumColumn )
			{
				$sumColumnIndexes[] = array_search( $sumColumn, $headers );
			}
		}
		else
		{
			$sumColumnIndexes[] = count( $headers ) - 1;
		}

		$rowGroupIndexes = $this->createGroupIndexes( $headers, $rowGroups );

		foreach ( $data as $index => $values )
		{
			$groupValues = array();
			foreach ( $rowGroupIndexes as $groupIndex )
			{
				$groupValues[ $groupIndex ] = $values[ $groupIndex ];
			}

			while ( count( $groupValues ) > 0 )
			{
				if ( ! $result->exists( $groupValues ) )
				{
					$totals = array_fill_keys( $sumColumnIndexes, 0 );
					$result->addValue( $groupValues, $totals );
				}

				$totals = $result->getValue( $groupValues );

				foreach ( $sumColumnIndexes as $columnIndex )
				{
					$totals[ $columnIndex ] += $values[ $columnIndex ];
				}

				$result->addValue( $groupValues, $totals );

				// Drop the last group
				array_pop( $groupValues );
			}
		}

		return $result;
	}

	/**
	 * Create the PivotTable xml.  The elements are defined in section 18.10
	 * Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf
	 * This only creates a basic pivot table based on data in a sheet
	 *
	 * Pivot tables are very complicated able to accommodate external sources like databases or OLAP servers with formatting,
	 * conditional filters, hierarchies, group, filters, formulas and more.
	 * None of this stuff is supported.  However, hopefully basic functionality provided combined with the available specification
	 * shows where code to support additional features can be added.
	 *
	 * @param Worksheet $sheet  			Sheet containing the source data
	 * @param string $dataRange				The range containing the source data
	 * @param string $definitionUniqueId	The unique id appended to create a unique file name of and reference id for the definition cache
	 * @param int $cacheId					Id held in workbooks
	 * @param array $sharedItems			An array of the fields and unique fields values
	 * @param string|int					The index of the sheet in which the pivot table will appear.
	 * 										This can be an exising or new sheet.
	 * @param int $colIndex					The left of the data area of the pivot table.
	 * @param int $rowIndex					The top of the data area of the pivot table.
	 * 										It will be an error if there is no space for paging columns.
	 * @param Groups $columnGroups			A list of the columns to display.  Can be empty.
	 * @param Groups $rowGroups				A list of the rows to display.  Can be empty.
	 * @param Groups $valueGroups			A list of the value columns
	 * @param string $dataCaption			The caption to display over the columns when they are displayed
	 * @return bool
	 * @throws Exception					If there are no row groups or there are conflicts between the row, columns and value groups
	 */
	private function createPivotTable( $sheet, $dataRange, $definitionUniqueId, $cacheId, $sharedItems, $sheetIndex, $colIndex, $rowIndex, $rowGroups = null, $columnGroups = null, $valueGroups = null, $name = "PivotTable1", $dataCaption = 'Data' )
	{
		// Two pivot table examples are shown below after the code.
		// For more information about the pivotTable element see section 18.10 of
		// Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf

		$sheet = $this->getSheetFromIndex( $sheetIndex );

		$boundary = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries( $dataRange );
		$leftCol = $boundary[0][0];
		$topRow = $boundary[0][1];

		// Access the data
		$data = $sheet->rangeToArray( $dataRange, null, false );
		$headers = array_shift( $data );

		$sharedItemsKeys = array_keys( $sharedItems );

		$first = $sharedItemsKeys[ 0 ];
		$last = $sharedItemsKeys[ count( $sharedItemsKeys ) - 1 ];

		if ( ! $rowGroups || ! $rowGroups->count() ) $rowGroups = new Groups( $first );
		if ( ! $valueGroups || ! $valueGroups->count() ) $valueGroups = new Groups( $last );
		if ( ! $columnGroups ) $columnGroups = new Groups();

		$overlap  = array_intersect( $rowGroups->getGroupNames(), $columnGroups->getGroupNames(), $valueGroups->getGroupNames() );
		if ( $overlap ) throw new \Exception('The group columns and the sum columns overlap: ' . join( ',', $overlap ) );

		$overlap = array_intersect( $headers, $rowGroups->getGroupNames() );
		if ( count( $overlap ) != $rowGroups->count() ) throw new \Exception('The row groups array contains invalid columns: ' . join( ',', $overlap ) );

		$overlap = array_intersect( $headers, $columnGroups->getGroupNames() );
		if ( count( $overlap ) != $columnGroups->count() ) throw new \Exception('The column groups array contains invalid columns: ' . join( ',', $overlap ) );

		$overlap = array_intersect( $headers, $valueGroups->getGroupNames() );
		if ( count( $overlap ) != $valueGroups->count() ) throw new \Exception('The value groups array contains invalid columns: ' . join( ',', $overlap ));

		// Compute the number of rows and columns.  This will include the number of
		// elements in the first column and the number of specified columns.
		$rowPivotGroups = $this->pivotGroups( $headers, $data, $rowGroups, $valueGroups );
		$columnPivotGroups = $this->pivotGroups( $headers, $data, $columnGroups, $valueGroups );
		$colCount = max( 1, $rowGroups->count() ) +		// The list of row groups - there will always be at least one
			count( $columnPivotGroups->getKeys() ) +	// The list of columns and sub-totals)
			max( 1, $valueGroups->count() );			// The 'grand' totals for the value groups
		$rowCount = 1 +									// Column selector buttons (or heading if no column groups)
			max( 1, $columnGroups->count() ) +			// The column (or row selector buttons)
			count( $rowPivotGroups->getKeys() ) +		// The list of row items and sub-totals
			1;											// Grand total

		$pivotTableRange =	Coordinate::stringFromColumnIndex( $colIndex ) . $rowIndex . ":" .
							Coordinate::stringFromColumnIndex( $colIndex + $colCount - 1 ) . ( $rowIndex + $rowCount - 1 );

        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
		$objWriter->startElement('pivotTableDefinition');
			$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
			$objWriter->writeAttribute('name', $name );
			$objWriter->writeAttribute('cacheId', $cacheId );
			// $objWriter->writeAttribute('dataOnRows', "1");
			$objWriter->writeAttribute('applyNumberFormats', "0" );
			$objWriter->writeAttribute('applyBorderFormats', "0" );
			$objWriter->writeAttribute('applyFontFormats', "0" );
			$objWriter->writeAttribute('applyPatternFormats', "0" );
			$objWriter->writeAttribute('applyAlignmentFormats', "0" );
			$objWriter->writeAttribute('applyWidthHeightFormats', "1" );
			$objWriter->writeAttribute('dataCaption', $dataCaption );
			$objWriter->writeAttribute('showMultipleLabel', "0" );
			$objWriter->writeAttribute('showMemberPropertyTips', "0" );
			$objWriter->writeAttribute('useAutoFormatting', "1" );
			$objWriter->writeAttribute('indent', "127" );
			$objWriter->writeAttribute('compact', "0" );
			$objWriter->writeAttribute('compactData', "0" );
			$objWriter->writeAttribute('gridDropZones', "1" );

			$objWriter->startElement('location');
				$objWriter->writeAttribute('ref', $pivotTableRange );
				$objWriter->writeAttribute('firstHeaderRow', "2" );
				$objWriter->writeAttribute('firstDataRow', "2" );
				$objWriter->writeAttribute('firstDataCol', "1" );
			$objWriter->endElement();

			$objWriter->startElement('pivotFields');
				$objWriter->writeAttribute('count', count( $sharedItemsKeys ) );

			if ( ! $rowGroups )
			{
				$rowGroups = new Groups();
			}

			foreach ( $sharedItems as $column => $items )
			{
				$objWriter->startElement('pivotField');
					$objWriter->writeAttribute('compact', "0" );
					$objWriter->writeAttribute('outline', "0" );
					$objWriter->writeAttribute('subtotalTop', "0" );
					$objWriter->writeAttribute('showAll', "0" );
					$objWriter->writeAttribute('includeNewItemsInFilter', "1" );

					$rowGroup = $rowGroups->getGroupByName( $column );
					$columnGroup = $columnGroups->getGroupByName( $column );

					if ( $rowGroup || $columnGroup )
					{
						$group = null;
						if ( $rowGroup )
						{
							$group = $rowGroup;
							$objWriter->writeAttribute('axis', "axisRow" );
							$objWriter->writeAttribute('sortType', $rowGroup->getSortType() );
						}
						else if ( $columnGroup )
						{
							$group = $columnGroup;
							$objWriter->writeAttribute('axis', "axisCol" );
							$objWriter->writeAttribute('sortType', $columnGroup->getSortType() );
						}

						$objWriter->startElement('items');
							$objWriter->writeAttribute('count', count( $items ) + 1 );

							foreach ( $items as $itemIndex => $item )
							{
								$objWriter->startElement('item');
									$objWriter->writeAttribute('x', $itemIndex );
									if ( $group && ! $group->getVisibleState( $item ) )
									{
										$objWriter->writeAttribute('h', 1 );
									}
								$objWriter->endElement();
							}

							$objWriter->startElement('item');
								$objWriter->writeAttribute('t', 'default');
							$objWriter->endElement();

						$objWriter->endElement();
					}
					else if ( $valueGroups->getGroupByName( $column ) )
					{
						$objWriter->writeAttribute('dataField', "1" );
					}

				$objWriter->endElement();
			}

			$objWriter->endElement();

			$objWriter->startElement('rowFields');
				$objWriter->writeAttribute('count', $rowGroups->count() );

				foreach ( $rowGroups as $row => $rowGroup )
				{
					$objWriter->startElement('field');
						$objWriter->writeAttribute('x', array_search( $row, $sharedItemsKeys ) );
					$objWriter->endElement();
				}
			$objWriter->endElement();

			$objWriter->startElement('rowItems');
				$objWriter->writeAttribute('count', count( $rowPivotGroups->getKeys() ) + 1 );

			/**
			 	The row items element has this kind of structure where each sub element of <i> is a member of respective
			 	row group in display column order.  When there is more than one column then there is a sub-total.  These
			 	rows are represented by the @t.
				<i>
					<x/>
					<x v="2"/>
				</i>
				<i t="default">
					<x/>
				</i>
				<i t="grand">
					<x/>
				</i>
			 */

			$rowGroupIndexes = $this->createGroupIndexes( $headers, $rowGroups );
			foreach ( $rowPivotGroups->getKeys() as $hash => $rowMembers )
			{
				$objWriter->startElement('i');
					if ( $rowGroups->count() > count( $rowMembers ) )
					{
						$objWriter->writeAttribute('t', 'default' );
					}

					$index = 0;
					foreach ( $rowMembers as $member )
					{
						$objWriter->startElement('x');
						$groupIndex = $rowGroupIndexes[ $index ];
						$group = $sharedItemsKeys[ $groupIndex ];
						$key = array_search( $member, $sharedItems[ $group ] );
						if ( $key )
						{
							$objWriter->writeAttribute('v', $key );
						}
						$objWriter->endElement();
						$index++;
					}

				$objWriter->endElement();
			}

				$objWriter->startElement('i');
					$objWriter->writeAttribute('t', 'grand' );

					$objWriter->startElement('x');
					$objWriter->endElement();

				$objWriter->endElement();

			$objWriter->endElement();

			if ( ! $columnGroups || ! $columnGroups->getGroupNames() )
			{
				if ( $valueGroups->count() > 1 )
				{
					/**
						<colFields count="1">
							<field x="-2"/>
						</colFields>
					 */
					$objWriter->startElement('colFields');
						$objWriter->writeAttribute('count', "1" );
						$objWriter->startElement('field');
							$objWriter->writeAttribute('x', "-2" ); // This means show values.  That is, the data fields.
						$objWriter->endElement();
					$objWriter->endElement();
				}
			}
			else
			{
				$objWriter->startElement('colFields');
					$objWriter->writeAttribute('count', $columnGroups->count() );
					foreach ( $columnGroups as $column => $columnmGroup )
					{
						$colIndex = array_search( $column, $sharedItemsKeys );
						$objWriter->startElement('field');
							$objWriter->writeAttribute('x', $colIndex ); // This means show values.  That is, the data fields.
						$objWriter->endElement();
					}
				$objWriter->endElement();
			}

			$objWriter->startElement('colItems');

			if ( ! $columnGroups || ! $columnGroups->count() )
			{
				$objWriter->writeAttribute('count', 1 );

				/**
					<colItems count="1">
						<i>
							<x/>
						</i>
					</colItems>
				 */
					$objWriter->startElement('i');

						$objWriter->startElement('x');
						$objWriter->endElement();

					$objWriter->endElement();
			}
			else
			{
				$objWriter->writeAttribute('count', count( $columnPivotGroups->getKeys() ) + 1 );

				/**
				 	The col items element has this kind of structure where each sub element of <i> is a member of respective
				 	column in display row group order.  When there is more than one column then there is a sub-total.  These
				 	columns are represented by the @t.
					<i>
						<x/>
						<x v="2"/>
					</i>
					<i t="default">
						<x/>
					</i>
					<i t="grand">
						<x/>
					</i>
				 */

				$columnIndexes = $this->createGroupIndexes( $headers, $columnGroups );
				foreach ( $columnPivotGroups->getKeys() as $hash => $columnMembers )
				{
					$objWriter->startElement('i');
						if ( $columnGroups->count() > count( $columnMembers ) )
						{
							$objWriter->writeAttribute('t', 'default' );
						}

						$index = 0;
						foreach ( $columnMembers as $member )
						{
							$objWriter->startElement('x');
							$columnIndex = $columnIndexes[ $index ];
							$column = $sharedItemsKeys[ $columnIndex ];
							$key = array_search( $member, $sharedItems[ $column ] );
							if ( $key )
							{
								$objWriter->writeAttribute('v', $key );
							}
							$objWriter->endElement();
							$index++;
						}

					$objWriter->endElement();
				}

					$objWriter->startElement('i');
						$objWriter->writeAttribute('t', 'grand' );

						$objWriter->startElement('x');
						$objWriter->endElement();

					$objWriter->endElement();

			}

			$objWriter->endElement();

			// $objWriter->startElement('pageFields');
			// $objWriter->endElement();

			$objWriter->startElement('dataFields');
				$objWriter->writeAttribute('count', $valueGroups->count() );

			// if ( $valueGroups )
			{
				foreach ( $valueGroups as $column => $valueGroup )
				{
					$columnIndex = array_search( $column, $headers );
					$objWriter->startElement('dataField');
						$objWriter->writeAttribute('name', 'Sum of ' . $column );
						$objWriter->writeAttribute('fld', $columnIndex );
						$objWriter->writeAttribute('baseField', 0 );
						$objWriter->writeAttribute('baseItem', 0 );
					$objWriter->endElement();
				}
			}
			// else
			// {
			// 	$sharedItems[ $last ];
            //
			// 	$objWriter->startElement('dataField');
			// 		$objWriter->writeAttribute('name', 'Sum of ' . $last );
			// 		$objWriter->writeAttribute('fld', '' );
			// 		$objWriter->writeAttribute('baseField', 0 );
			// 		$objWriter->writeAttribute('baseItem', 0 );
			// 	$objWriter->endElement();
			// }

			$objWriter->endElement();

			// $objWriter->startElement('conditionalFormats');
			// $objWriter->endElement();

			$objWriter->startElement('pivotTableStyleInfo');
				$objWriter->writeAttribute('showRowHeaders', "1" );
				$objWriter->writeAttribute('showColHeaders', "1" );
				$objWriter->writeAttribute('showRowStripes', "0" );
				$objWriter->writeAttribute('showColStripes', "0" );
				$objWriter->writeAttribute('showLastColumn', "1" );

			$objWriter->endElement();

		$objWriter->endElement();

		$zip = new ZipArchiveX();
		$xml = $objWriter->getData();
		// error_log( $xml );
		$xml = $zip->formattedXml( $xml );
		// echo $xml;
		unset( $zip );

		$pivotTable = $this->addPivotTable( "rId$definitionUniqueId", "xl/pivotTables/pivotTable$definitionUniqueId.xml", $xml, $cacheId, "xl/worksheets/sheet" . ( $this->getIndex( $sheet ) + 1 ) . ".xml" );

		$rId = $this->pivotCacheDefinitionCollection->getPivotCacheIndex( $cacheId );
		$path = $this->pivotCacheDefinitionCollection->getPivotCacheDefinitionPath( $rId );
		$pivotTable->cacheDefinitionPath = $path;

		return true;

		/*
		 	Both examples are based on the same data cache.  The diffence is in the field selection and filtering.

		 	The first example is the default layout generated by Excel for the data.  That is, it sums the data field
		 	(the last numeric one) by the first field (Account).

		 	The second is modified to show two summed columns grouped by a different field (Genre) and has a page
		 	filter (Account)

			<pivotTableDefinition>
				<location/>
				<pivotFields/>
				<rowFields/>
				<rowItems/>
				<colFields/>
				<colItems/>
				<pageFields/>
				<dataFields/>
				<conditionalFormats/>
				<pivotTableStyleInfo/>
			</pivotTableDefinition>

			---------------------------
			| Sum of Size |           |
			---------------------------
			| Account    V|	Total     |
			---------------------------
			| Megan       |     72000 |
			| Hannah      |     83000 |
			| Vicky       |     42000 |
			| Ian         |     92000 |
			| Michael     |     72000 |
			| Daniel      |     85000 |
			---------------------------
			| Grand Total |    446000 |
			---------------------------

			<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<pivotTableDefinition
				xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
				name="PivotTable1"
				cacheId="4" dataOnRows="1"
				applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0"
				applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Data"
				showMultipleLabel="0" showMemberPropertyTips="0" useAutoFormatting="1" indent="127"
				compact="0" compactData="0" gridDropZones="1"
			>
				<location ref="B13:C21"
					<!-- Specifies the first row of  the PivotTable header, relative to the top left cell in the ref value. -->
					firstHeaderRow="2" <-- Controls where the dropdown button appears  relative to the location
					<!-- Specifies the first row of the PivotTable data, relative to the top left cell in the ref value. -->
					firstDataRow="2"
					<!-- Specifies the first column of the PivotTable data, relative to the top left cell in the ref value ->>
					firstDataCol="1" <-- Don't seem to have an effect but if these two are changed the workbook fails to open
				/>
				<pivotFields count="5">
					<pivotField axis="axisRow" compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1">
						<items count="7">
							<item x="0"/>
							<item x="1"/>
							<item x="2"/>
							<item x="3"/>
							<item x="4"/>
							<item x="5"/>
							<item t="default"/>
						</items>
					</pivotField>
					<pivotField compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1"/>
					<pivotField compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1"/>
					<pivotField compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1"/>
					<pivotField dataField="1" compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1"/>
				</pivotFields>
				<rowFields count="1">
					<field x="0"/>
				</rowFields>
				<rowItems count="7">
					<i>
						<x/>
					</i>
					<i>
						<x v="1"/>
					</i>
					<i>
						<x v="2"/>
					</i>
					<i>
						<x v="3"/>
					</i>
					<i>
						<x v="4"/>
					</i>
					<i>
						<x v="5"/>
					</i>
					<i t="grand">
						<x/>
					</i>
				</rowItems>
				<colItems count="1">
					<i/>
				</colItems>
				<dataFields count="1">
					<dataField name="Sum of Total Size" fld="4" baseField="0" baseItem="0"/>
				</dataFields>
				<pivotTableStyleInfo showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>
			</pivotTableDefinition>

			----------------------------------------------
			| Account     | All          |               |
			----------------------------------------------
			|             | Data         |               |
			| Genre      V| Sum of Size  | Sum of Images |
			----------------------------------------------
			| Floral      |        42000 |            25 |
			| Landscapes  |        83000 |            31 |
			| Portraits   |        72000 |            20 |
			----------------------------------------------
			| Grand Total |	      197000 |            76 |
			----------------------------------------------

			<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<pivotTableDefinition
			    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
			    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr"
			    xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
			    xr:uid="{00000000-0007-0000-0000-000000000000}" name="PivotTable1" cacheId="0"
			    applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0"
			    applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Data" updatedVersion="6"
			    showMultipleLabel="0" showMemberPropertyTips="0" useAutoFormatting="1" indent="127" compact="0" compactData="0" gridDropZones="1">
			    <location ref="B13:D18" firstHeaderRow="1" firstDataRow="2" firstDataCol="1" rowPageCount="1" colPageCount="1"/>
			    <pivotFields count="5">
			        <pivotField axis="axisPage" compact="0" outline="0" subtotalTop="0" multipleItemSelectionAllowed="1" showAll="0" includeNewItemsInFilter="1">
			            <items count="7">
			                <item x="0"/>
			                <item x="1"/>
			                <item x="2"/>
			                <item h="1" x="3"/>
			                <item h="1" x="4"/>
			                <item h="1" x="5"/>
			                <item t="default"/>
			            </items>
			        </pivotField>
			        <pivotField axis="axisRow" compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1">
			            <items count="4">
			                <item x="2"/>
			                <item x="1"/>
			                <item x="0"/>
			                <item t="default"/>
			            </items>
			        </pivotField>
			        <pivotField dataField="1" compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1"/>
			        <pivotField compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1"/>
			        <pivotField dataField="1" compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1"/>
			    </pivotFields>
			    <rowFields count="1">
			        <field x="1"/>
			    </rowFields>
			    <rowItems count="4">
			        <i>
			            <x/>
			        </i>
			        <i>
			            <x v="1"/>
			        </i>
			        <i>
			            <x v="2"/>
			        </i>
			        <i t="grand">
			            <x/>
			        </i>
			    </rowItems>
			    <colFields count="1">
			        <field x="-2"/>
			    </colFields>
			    <colItems count="2">
			        <i>
			            <x/>
			        </i>
			        <i i="1">
			            <x v="1"/>
			        </i>
			    </colItems>
			    <pageFields count="1">
			        <pageField fld="0" hier="-1"/>
			    </pageFields>
			    <dataFields count="2">
			        <dataField name="Sum of Total Size" fld="4" baseField="0" baseItem="0"/>
			        <dataField name="Sum of Images" fld="2" baseField="0" baseItem="0"/>
			    </dataFields>
			    <pivotTableStyleInfo showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>
			    <extLst>
			        <ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}"
			            xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout">
			            <xpdl:pivotTableDefinition16 EnabledSubtotalsDefault="0" SubtotalsOnTopDefault="0"/>
			        </ext>
			    </extLst>
			</pivotTableDefinition>
		 */
	}
}
