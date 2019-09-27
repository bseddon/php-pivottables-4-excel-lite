<?php

namespace lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Writer\Xlsx;

use PhpOffice\PhpSpreadsheet\Shared\XMLWriter;
use lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing;
use PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;
use PhpOffice;
// use PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart;
use lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheDefinition;
use lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotTable;
use lyquidity\xbrl_validate\PhpOffice\PhpSpreadsheet\Xlsx\PivotCacheRecords;

require_once __DIR__ . "/../Spreadsheet.php";

class Rels extends PhpOffice\PhpSpreadsheet\Writer\Xlsx\Rels
{
	/**
     * Added to support writing relationships for cache files (to record set files)
     * @param Spreadsheet $spreadsheet
     * @param PivotCacheDefinition $definition
     * @return string
     */
	public function writeCacheRelationships(Spreadsheet $spreadsheet, PivotCacheDefinition $definition)
	{
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        $i = 1;
		foreach( $definition->records as $path => /** @var PivotCacheRecords $records */ $records )
		{
			$this->writeRelationship(
				$objWriter,
				$i,
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords',
				basename( $records->path )
			);
			$i++;
		}

        $objWriter->endElement();

        return $objWriter->getData();
	}

	/**
	 * Added to record pivot table relationships (to cache definitions)
	 * @param Spreadsheet $spreadsheet
	 * @param PivotTable $table
	 * @return string
	 */
	public function writePivotTableRelationships(Spreadsheet $spreadsheet, PivotTable $table)
	{
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $result = $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $result = $objWriter->startElement('Relationships');
        $result = $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

		$this->writeRelationship(
			$objWriter,
			1,
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
			preg_replace("!^xl/!", "../", $table->cacheDefinitionPath )
		);

		$objWriter->endElement();

		return $objWriter->getData();
	}

	/**
     * Write workbook relationships to XML format.
     * BMS Overridden to add cache definition relationships
     *
     * @param Spreadsheet $spreadsheet
     *
     * @throws WriterException
     *
     * @return string XML Output
     */
    public function writeWorkbookRelationships(\PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet)
    {
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Relationship styles.xml
        $this->writeRelationship(
            $objWriter,
            1,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
            'styles.xml'
        );

        // Relationship theme/theme1.xml
        $this->writeRelationship(
            $objWriter,
            2,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
            'theme/theme1.xml'
        );

        // Relationship sharedStrings.xml
        $this->writeRelationship(
            $objWriter,
            3,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
            'sharedStrings.xml'
        );

        // Relationships with sheets
        $sheetCount = $spreadsheet->getSheetCount();
        for ($i = 0; $i < $sheetCount; ++$i) {
            $this->writeRelationship(
                $objWriter,
                ($i + 1 + 3),
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                'worksheets/sheet' . ($i + 1) . '.xml'
            );
        }
        // Relationships for vbaProject if needed
        // id : just after the last sheet
        if ($spreadsheet->hasMacros()) {
            $this->writeRelationShip(
                $objWriter,
                ($i + 1 + 3),
                'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
                'vbaProject.bin'
            );
            ++$i; //increment i if needed for an another relation
        }

        if ( $spreadsheet->pivotCacheDefinitionCollection->hasPivotCacheDefinitions() )
        {
        	foreach ( $spreadsheet->pivotCacheDefinitionCollection as $path => /** @var PivotCacheDefinition $definition */ $definition )
        	{
	            $this->writeRelationShip(
	                $objWriter,
	                ($i + 1 + 3),
	                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
	                preg_replace("!^xl/!","", $path )
	            );
	            $definition->afterSaveReferenceId = "rId" . ( $i + 1 + 3 );
	            ++$i; //increment i if needed for an another relation

        	}
        }

        $objWriter->endElement();

        return $objWriter->getData();
    }

    /**
     * Write worksheet relationships to XML format.
     * Overidden to add relationship to the sheets (to pivot tables)
     *
     * Numbering is as follows:
     *     rId1                 - Drawings
     *  rId_hyperlink_x     - Hyperlinks
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pWorksheet
     * @param int $pWorksheetId
     * @param bool $includeCharts Flag indicating if we should write charts
     *
     * @throws WriterException
     *
     * @return string XML Output
     */
    public function writeWorksheetRelationships(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pWorksheet, $pWorksheetId = 1, $includeCharts = false)
    {
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Write drawing relationships?
        $d = 0;
        $drawingOriginalIds = [];
        $unparsedLoadedData = $pWorksheet->getParent()->getUnparsedLoadedData();
        if (isset($unparsedLoadedData['sheets'][$pWorksheet->getCodeName()]['drawingOriginalIds'])) {
            $drawingOriginalIds = $unparsedLoadedData['sheets'][$pWorksheet->getCodeName()]['drawingOriginalIds'];
        }

        if ($includeCharts) {
            $charts = $pWorksheet->getChartCollection();
        } else {
            $charts = [];
        }

        if (($pWorksheet->getDrawingCollection()->count() > 0) || (count($charts) > 0) || $drawingOriginalIds) {
            $relPath = '../drawings/drawing' . $pWorksheetId . '.xml';
            $rId = ++$d;

            if (isset($drawingOriginalIds[$relPath])) {
                $rId = (int) (substr($drawingOriginalIds[$relPath], 3));
            }

            $this->writeRelationship(
                $objWriter,
                $rId,
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                $relPath
            );
        }

        // Write hyperlink relationships?
        $i = 1;
        foreach ($pWorksheet->getHyperlinkCollection() as $hyperlink) {
            if (!$hyperlink->isInternal()) {
                $this->writeRelationship(
                    $objWriter,
                    '_hyperlink_' . $i,
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                    $hyperlink->getUrl(),
                    'External'
                );

                ++$i;
            }
        }

		/**
		 *
		 * @var Spreadsheet $spreadsheet
		 */
		$spreadsheet = $pWorksheet->getParent();
		if ( $spreadsheet->pivotTables->hasPivotTables() ) // Are there any pivot tables?
		{
			// If so, is there a pivot table for this sheet?
			// if ( ( $pivotTables = $spreadsheet->pivotTables->ownedBy("xl/worksheets/sheet$pWorksheetId.xml") ) )
			if ( ( $pivotTables = $spreadsheet->pivotTables->ownedBy( $pWorksheet->getCodeName() ) ) )
			{
				foreach ( $pivotTables as $path => /** @var PivotTable $pivotTable */ $pivotTable )
				{
					$this->writeRelationship(
						$objWriter,
						$i,
						'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
		                preg_replace( "!^xl/!", "../", $pivotTable->path )
		            );
					$pivotTable->afterSaveReferenceId = $spreadsheet->pivotCacheDefinitionCollection->getPivotCacheDefinitionByPath( $pivotTable->cacheDefinitionPath )->afterSaveReferenceId;
					$i++;
				}
			}
        }

        // Write comments relationship?
        $i = 1;
        if (count($pWorksheet->getComments()) > 0) {
            $this->writeRelationship(
                $objWriter,
                '_comments_vml' . $i,
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
                '../drawings/vmlDrawing' . $pWorksheetId . '.vml'
            );

            $this->writeRelationship(
                $objWriter,
                '_comments' . $i,
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
                '../comments' . $pWorksheetId . '.xml'
            );
        }

        // Write header/footer relationship?
        $i = 1;
        if (count($pWorksheet->getHeaderFooter()->getImages()) > 0) {
            $this->writeRelationship(
                $objWriter,
                '_headerfooter_vml' . $i,
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
                '../drawings/vmlDrawingHF' . $pWorksheetId . '.vml'
            );
        }

        $this->writeUnparsedRelationship($pWorksheet, $objWriter, 'ctrlProps', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp');
        $this->writeUnparsedRelationship($pWorksheet, $objWriter, 'vmlDrawings', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing');
        $this->writeUnparsedRelationship($pWorksheet, $objWriter, 'printerSettings', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings');

        $objWriter->endElement();

        return $objWriter->getData();
    }

    /**
     * Included because the parent function is private
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pWorksheet
     * @param XMLWriter $objWriter
     * @param unknown $relationship
     * @param unknown $type
     */
    private function writeUnparsedRelationship(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pWorksheet, XMLWriter $objWriter, $relationship, $type)
    {
        $unparsedLoadedData = $pWorksheet->getParent()->getUnparsedLoadedData();
        if (!isset($unparsedLoadedData['sheets'][$pWorksheet->getCodeName()][$relationship])) {
            return;
        }

        foreach ($unparsedLoadedData['sheets'][$pWorksheet->getCodeName()][$relationship] as $rId => $value) {
            $this->writeRelationship(
                $objWriter,
                $rId,
                $type,
                $value['relFilePath']
            );
        }
    }

    /**
     * Write Override content type.
     * Included here because the method in the parent is private
     *
     * @param XMLWriter $objWriter XML Writer
     * @param int $pId Relationship ID. rId will be prepended!
     * @param string $pType Relationship type
     * @param string $pTarget Relationship target
     * @param string $pTargetMode Relationship target mode
     *
     * @throws WriterException
     */
    private function writeRelationship(XMLWriter $objWriter, $pId, $pType, $pTarget, $pTargetMode = '')
    {
        if ($pType != '' && $pTarget != '')
        {
            // Write relationship
            $result = $objWriter->startElement('Relationship');
            $result = $objWriter->writeAttribute('Id', 'rId' . $pId);
            $result = $objWriter->writeAttribute('Type', $pType);
            $result = $objWriter->writeAttribute('Target', $pTarget);

            if ($pTargetMode != '')
            {
                $result = $objWriter->writeAttribute('TargetMode', $pTargetMode);
            }

            $result = $objWriter->endElement();
        } else {
            throw new WriterException('Invalid parameters passed.');
        }
    }
}
