<?php

namespace lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx;

use lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheRecords;

/**
 * A collection of PivotCacheDefinition instances
 */
class PivotCacheRecordsCollection implements \IteratorAggregate
{
	/**
	 * Path/Xml for references in pivotCachedefinitionsX.xml.rels
	 * @var array
	 */
	private $pivotCacheRecords = [];
	/**
	 * rId/Path  for references in pivotCachedefinitions[x].xml.rels
	 * @var array
	 */
	private $pivotCacheRecordIndex = [];

	public function __construct( $records = null )
	{
		if ( $records instanceof PivotCacheRecords )
		{
			$this->addPivotCacheRecords( $records );
		}
		else if ( is_array( $records ) || $records instanceof PivotCacheRecordsCollection )
		{
			foreach ( $records as $_records )
			{
				$this->addPivotCacheRecords( $_records );
			}
		}
	}

	public function getIterator()
	{
		return (function ()
		{
			reset($this->pivotCacheRecords);
			// while(list($key, $val) = each($this->pivotCacheRecords))
			foreach ( $this->pivotCacheRecords as $key => $val )
			{
				yield $key => $val;
			}
		})();
	}

	// A collection of functions to record pivot cache record references

	/**
	 * Add a records instance to the collection
	 * @param PivotCacheRecords $records
	 */
	public function addPivotCacheRecords( $records )
	{
		$this->pivotCacheRecords[ $records->path ] = $records;
		$this->pivotCacheRecordIndex[ $records->referenceId ] = $records->path;
	}

	/**
	 * Return the number of records in the collection
	 * @return number
	 */
	public function hasPivotCacheRecords()
	{
		return count( $this->pivotCacheRecords );
	}

	/**
	 * Return the records instance with the zip path
	 * @param string $path
	 * @return NULL|PivotCacheRecords
	 */
	public function getPivotCacheRecordByPath( $path )
	{
		if ( ! isset( $this->pivotCacheRecords[ $path ] ) ) return null;
		return $this->pivotCacheRecords[ $path ] ;
	}

	/**
	 * Return the records with the reference id
	 * @param string $rId
	 * @return NULL|PivotCacheRecords
	 */
	public function getPivotCacheRecordsById( $rId )
	{
		if ( ! isset( $this->pivotCacheRecordIndex[ $rId ] ) ) return null;
		$path = $this->pivotCacheRecordIndex[ $rId ];

		if ( ! isset( $this->pivotCacheRecords[ $path ] ) ) return null;
		return $this->pivotCacheRecords[ $path ] ;
	}

}
