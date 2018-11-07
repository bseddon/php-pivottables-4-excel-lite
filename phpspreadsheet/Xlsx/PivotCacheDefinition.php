<?php

namespace lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx;

/**
 * Class to represent a cache definition
 */
class PivotCacheDefinition
{
	/**
	 * The workbook reference id for this definition
	 * @var string
	 */
	public $referenceId = "";
	/**
	 * The zip file path to this definition
	 * @var string
	 */
	public $path = "";
	/**
	 * The workbook cache id associated with this cache definition
	 * @var string
	 */
	private $cacheId = null;
	/**
	 * The raw xml from the cache definition file
	 * @var string
	 */
	public $xml = null;
	/**
	 * Records the cache records assoicated with this cache definition
	 * @var PivotCacheRecordsCollection
	 */
	public $records = null;
	/**
	 * The referenceId *after* the defintion has been written as a relation
	 * @var unknown
	 */
	public $afterSaveReferenceId;

	static $lastCacheId = 0;

	/**
	 * Constructor
	 * @param string $path
	 * @param string $xml
	 * @param string $referenceId
	 * @param string|null $cacheId
	 */
	public function __construct( $path, $xml, $referenceId, $cacheId = null )
	{
		$this->records = new PivotCacheRecordsCollection();

		$this->referenceId = $referenceId;
		$this->path = $path;
		$this->xml = $xml;

		if ( ! $cacheId ) return;
		$this->addCacheId( $cacheId );
	}

	/**
	 * Add the cache id that refers to this definition
	 * @param string $cacheId
	 */
	public function addCacheId( $cacheId )
	{
		$this->cacheId = $cacheId;
		if ( PivotCacheDefinition::$lastCacheId >= $cacheId ) return;

		PivotCacheDefinition::$lastCacheId = $cacheId;
	}

	/**
	 * Gets the stored cache id
	 * @return string
	 */
	public function getCacheId()
	{
		return $this->cacheId;
	}

	public static function getNextCacheId()
	{
		return PivotCacheDefinition::$lastCacheId + 1;
	}

	/**
	 * Add the records belonging to this cache definition
	 * @param unknown $recordsCollection
	 * @throws \ArgumentException
	 */
	public function addCacheRecords( $recordsCollection )
	{
		if ( ! $recordsCollection instanceof PivotCacheRecordsCollection )
		{
			throw new \ArgumentException("Invalid argument type.  Expected PivotCacheRecordsCollection");
		}

		if ( ! $this->records || ! $this->records->hasPivotCacheRecords() )
		{
			$this->records = $recordsCollection;
		}
		else
		{
			foreach( $recordsCollection as $records )
			{
				$this->records->addPivotCacheRecords( $records );
			}
		}
	}
}