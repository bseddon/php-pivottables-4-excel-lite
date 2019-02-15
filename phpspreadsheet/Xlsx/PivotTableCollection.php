<?php

namespace lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx;

use lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotTable;

/**
 * A collection of PivotCacheDefinition instances
 */
class PivotTableCollection implements \IteratorAggregate
{
	/**
	 * Path/Xml for references in sheet[x].xml.rels
	 * @var array
	 */
	private $pivotTables = [];
	/**
	 * rId/Path  for references in sheet[x].xml.rels
	 * @var array
	 */
	private $pivotTableIndex = [];

	/**
	 *
	 * @param PivotTable|PivotTableCollection|array $pivotTable
	 */
	public function __construct( $pivotTable = null )
	{
		if ( $pivotTable instanceof PivotTable )
		{
			$this->addPivotTable( $pivotTable );
		}
		else if ( is_array( $pivotTable ) || $pivotTable instanceof PivotTableCollection )
		{
			foreach ( $pivotTable as $_pivotTable )
			{
				$this->addPivotTable( $_pivotTable );
			}
		}
	}

	public function getIterator()
	{
		return (function ()
		{
			reset($this->pivotTables);
			// while(list($key, $val) = each($this->pivotTables))
			foreach ( $this->pivotTables as $key => $val )
			{
				yield $key => $val;
			}
		})();
	}

	// A collection of functions to record pivot cache record references

	/**
	 * Add a records instance to the collection
	 * @param PivotTable $pivotTable
	 */
	public function addPivotTable( $pivotTable )
	{
		$this->pivotTables[ $pivotTable->path ] = $pivotTable;
		$this->pivotTable[ $pivotTable->referenceId ] = $pivotTable->path;
	}

	/**
	 * Return the number of tables in the collection
	 * @return number
	 */
	public function hasPivotTables()
	{
		return count( $this->pivotTables );
	}

	/**
	 * Return the table instance with the zip path
	 * @param string $path
	 * @return NULL|PivotTable
	 */
	public function getPivotTableByPath( $path )
	{
		if ( ! isset( $this->pivotTables[ $path ] ) ) return null;
		return $this->pivotTables[ $path ] ;
	}

	/**
	 * Return the table with the reference id
	 * @param string $rId
	 * @return NULL|PivotTable
	 */
	public function getPivotTableById( $rId )
	{
		if ( ! isset( $this->pivotTableIndex[ $rId ] ) ) return null;
		$path = $this->pivotTableIndex[ $rId ];

		if ( ! isset( $this->pivotTables[ $path ] ) ) return null;
		return $this->pivotTables[ $path ] ;
	}

	public function ownedBy( $path )
	{
		return array_filter(
			$this->pivotTables,
			function( $table ) use( $path )
			{
				return $table->sheetName == $path;
			}
		);
	}
}
