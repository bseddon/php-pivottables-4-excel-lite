<?php

namespace lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx;

use lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheDefinition;

/**
 * A collection of PivotCacheDefinition instances
 */
class PivotCacheDefinitionCollection implements \IteratorAggregate
{
	/**
	 * Path/Xml for references in workbook.xml.rels
	 * @var PivotCacheDefinition[]
	 */
	private $pivotCacheDefinitions = [];
	/**
	 * rId/Path for references in workbook.xml.rels
	 * @var array
	 */
	private $pivotCacheDefinitionIndex = [];
	/**
	 * cacheId/rId for references in workbook.xml.rels
	 * @var array
	 */
	private $pivotCaches = [];

	/**
	 * Implements the IteratorAggregator interface member
	 * {@inheritDoc}
	 * @see IteratorAggregate::getIterator()
	 */
	public function getIterator()
	{
		return (function ()
		{
			reset($this->pivotCacheDefinitions);
			// while(list($key, $val) = each($this->pivotCacheDefinitions))
			foreach ( $this->pivotCacheDefinitions as $key => $val )
			{
				yield $key => $val;
			}
		})();
	}

	/**
	 * Add a definition to the collection
	 * @param PivotCacheDefinition $definition
	 */
	public function addPivotCacheDefinition( $definition )
	{
		$this->pivotCacheDefinitions[ $definition->path ] = $definition;
		$this->pivotCacheDefinitionIndex[ $definition->referenceId ] = $definition->path;
		if ( ! $definition->getCacheId() ) return;

		$this->pivotCaches[ $definition->cacheId ] = $definition->referenceId;
	}

	/**
	 * Returns true if there are existing definitions
	 * @return bool
	 */
	public function hasPivotCacheDefinitions()
	{
		return count( $this->pivotCacheDefinitions );
	}

	/**
	 * Allow a caller to retrieve a defintion by its path
	 * @param string $path
	 * @return NULL|PivotCacheDefinition
	 */
	public function getPivotCacheDefinitionByPath( $path )
	{
		if ( ! isset( $this->pivotCacheDefinitions[ $path ] ) ) return null;
		return $this->pivotCacheDefinitions[ $path ] ;
	}

	/**
	 * Allow a caller to retrieve a defintion by its reference id
	 * @param string $path
	 * @return NULL|PivotCacheDefinition
	 */
	public function getPivotCacheDefinitionById( $rId )
	{
		if ( ! isset( $this->pivotCacheDefinitionIndex[ $rId ] ) ) return null;
		$path = $this->pivotCacheDefinitionIndex[ $rId ];

		if ( ! isset( $this->pivotCacheDefinitions[ $path ] ) ) return null;
		return $this->pivotCacheDefinitions[ $path ] ;
	}

	/**
	 * Allow a caller to retrieve a path by its reference id
	 * @param string $path
	 * @return NULL|PivotCacheDefinition
	 */
	public function getPivotCacheDefinitionPath( $rId )
	{
		if ( ! isset( $this->pivotCacheDefinitionIndex[ $rId ] ) ) return null;
		return $this->pivotCacheDefinitionIndex[ $rId ];
	}

	// A collection of functions to index workbook pivot cache references

	/**
	 * Add a reference id for a workbook cache id
	 * @param string $cacheId
	 * @param stirng $rId
	 * @return void
	 * @throws Exception If the $cacheId is not numeric
	 */
	public function addPivotCacheIndex( $cacheId, $rId )
	{
		if ( ! is_numeric( $cacheId ) ) throw new \Exception();

		$this->pivotCaches[ $cacheId ] = $rId;
		$definition = $this->getPivotCacheDefinitionById( $rId );
		$definition->addCacheId( $cacheId );
	}

	/**
	 * Return a reference id for a cache id.
	 * @param unknown $cacheId
	 * @return NULL|string
	 */
	public function getPivotCacheIndex( $cacheId )
	{
		if ( ! isset( $this->pivotCaches[ $cacheId ] ) ) return null;
		return $this->pivotCaches[ $cacheId ];
	}

	/**
	 *
	 * @param unknown $cacheId
	 * @return NULL|PivotCacheDefinition
	 */
	public function getPivotCache( $cacheId )
	{
		if ( ! isset( $this->pivotCaches[ $cacheId ] ) ) return null;
		$rId = $this->pivotCaches[ $cacheId ];

		return $this->getPivotCacheDefinitionById( $rId );
	}

	/**
	 * Return an array of all the cacheId recorded
	 * @return array;
	 */
	public function getCacheIds()
	{
		return array_keys( $this->pivotCaches );
	}

}
