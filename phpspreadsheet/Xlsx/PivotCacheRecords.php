<?php

namespace lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx;

/**
 * Class to represent a cache records set
 */
class PivotCacheRecords
{
	/**
	 * Workbook reference id for this record set
	 * @var string
	 */
	public $referenceId = "";
	/**
	 * Zip file path to this record set
	 * @var string
	 */
	public $path = "";
	/**
	 * Raw xml of the cache definition
	 * @var string
	 */
	public $xml = null;
	/**
	 * Zip file path of the 'owing' cache definition
	 * @var string
	 */
	public $cacheDefinitionPath = null;

	public function __construct( $path, $xml, $referenceId, $cacheDefinitionPath = null )
	{
		$this->referenceId = $referenceId;
		$this->path = $path;
		$this->xml = $xml;
		$this->cacheDefinitionPath = $cacheDefinitionPath;
	}
}