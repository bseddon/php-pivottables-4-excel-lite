<?php

namespace lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx;

/**
 * Class to represent a cache records set
 */
class PivotTable
{
	/**
	 * Workbook reference id for this able
	 * @var string
	 */
	public $referenceId = "";
	/**
	 * Zip file path to the table file
	 * @var string
	 */
	public $path = "";
	/**
	 * Raw xml of the table definition
	 * @var string
	 */
	public $xml = null;
	/**
	 * ame of the 'owing' sheet
	 * @var string
	 */
	public $sheetName = null;

	/**
	 * The cache id of the pivot table
	 * @var string
	 */
	public $cacheId;
	/**
	 * The referenceId *after* the defintion has been written as a relation
	 * @var unknown
	 */
	public $afterSaveReferenceId;
	public $cacheDefinitionPath;

	public function __construct( $path, $xml, $referenceId, $cacheId, $sheetName = null )
	{
		$this->referenceId = $referenceId;
		$this->path = $path;
		$this->xml = $xml;
		$this->cacheId = $cacheId;
		$this->sheetName = $sheetName;
	}

}