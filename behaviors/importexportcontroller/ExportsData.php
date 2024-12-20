<?php

namespace SKasianov\ExcelImportExport\Behaviors\ImportExportController;

use ApplicationException;
use Backend\Behaviors\ListController;
use Backend\Widgets\Lists;
use October\Rain\Support\Facades\File;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Response;

trait ExportsData
{
    protected function makeExportFileName($mode = 'json'): string
    {
        // Locate filename
        /** @noinspection PhpPossiblePolymorphicInvocationInspection */
        $fileName = $this->controller->importExportGetFileName();
        if (! $fileName) {
            $fileName = $this->getConfig('export[fileName]', 'export');
        }

        // Remove extension
        $fileName = File::name($fileName);

        return $fileName.'.'.$mode;
    }

    /**
     * @throws ApplicationException
     */
    protected function checkUseListExportMode(): \Illuminate\Http\Response|bool
    {
        if (! $useList = $this->getConfig('export[useList]')) {
            return false;
        }

        if (! $this->controller->isClassExtendedWith(ListController::class)) {
            throw new ApplicationException(__("You must implement the controller behavior ListController with the export 'useList' option enabled."));
        }

        if (is_array($useList)) {
            $listDefinition = array_get($useList, 'definition');
        } else {
            $listDefinition = $useList;
        }

        return $this->exportFromList($listDefinition, $this->getConfig('export'));
    }

    /**
     * @throws ApplicationException
     */
    public function exportFromList($definition = null, $options = []): \Illuminate\Http\Response
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $lists = $this->controller->makeLists();
        $widget = $lists[$definition] ?? reset($lists);

        $options = array_merge([
            'fileFormat' => $this->getConfig('defaultFormatOptions[fileFormat]', 'csv'),
            'delimiter' => $this->getConfig('defaultFormatOptions[delimiter]', ','),
            'enclosure' => $this->getConfig('defaultFormatOptions[enclosure]', '"'),
            'escape' => $this->getConfig('defaultFormatOptions[escape]', '\\'),
            'encoding' => $this->getConfig('defaultFormatOptions[encoding]', 'utf-8'),
        ], $options);

        $fileFormat = $options['fileFormat'];
        $filename = e($this->makeExportFileName($fileFormat));

        if ($fileFormat === 'json') {
            return Response::make(
                $this->exportFromListAsJson($widget, $options),
                200,
                [
                    'Content-Type' => 'application/json',
                    'Content-Disposition' => sprintf('%s; filename="%s"', 'attachment', $filename),
                ]
            );
        }

        if ($fileFormat === 'csv') {
            return Response::make(
                $this->exportFromListAsCsv($widget, $options),
                200,
                [
                    'Content-Type' => 'text/csv',
                    'Content-Transfer-Encoding' => 'binary',
                    'Content-Disposition' => sprintf('%s; filename="%s"', 'attachment', $filename),
                ]
            );
        }

        if ($fileFormat === 'xlsx') {
            $this->exportFromListAsSpreadsheet($widget, $options);
        }

        if ($fileFormat === 'xls') {
            $this->exportFromListAsSpreadsheet($widget, $options);
        }

        if ($fileFormat === 'ods') {
            $this->exportFromListAsSpreadsheet($widget, $options);
        }

        throw new ApplicationException('Unsupported file format');
    }

    /**
     * @throws ApplicationException
     */
    private function exportFromListAsSpreadsheet(Lists $list, array $options): void
    {
        // Locate columns from widget
        $columns = $list->getVisibleColumns();

        // Parse options
        $options = array_merge([
            'firstRowTitles' => true,
            'savePath' => null,
            'useOutput' => false,
            'fileName' => null,
        ], $options);

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $alphabet = range('A', 'Z');

        $rowNum = 1;
        if ($options['firstRowTitles']) {
            $colNum = 0;
            foreach ($columns as $column) {
                $letter = $alphabet[$colNum];
                $sheet->setCellValue("$letter$rowNum", $column->short_label ?? $column->label);
                $colNum++;
            }

            $rowNum = 2;
        }

        // Add records
        $getter = $this->getConfig('export[useList][raw]', false)
            ? 'getColumnValueRaw'
            : 'getColumnValue';

        $query = $list->prepareQuery();
        $results = $query->get();

        if ($event = $list->fireSystemEvent('backend.list.extendRecords', [&$results])) {
            $results = $event;
        }

        foreach ($results as $record) {
            $colNum = 0;
            foreach ($columns as $column) {
                $letter = $alphabet[$colNum];
                $value = $list->$getter($record, $column);
                if (is_array($value)) {
                    $value = implode('|', $value);
                }
                $sheet->setCellValue("$letter$rowNum", $value);
                $colNum++;
            }
            $rowNum++;
        }

        $writer = match ($options['fileFormat']) {
            'xlsx' => new Xlsx($spreadsheet),
            'xls' => new Xls($spreadsheet),
            'ods' => new Ods($spreadsheet),
            default => throw new ApplicationException('Unsupported file format'),
        };

        match ($options['fileFormat']) {
            'xlsx' => header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
            'xls' => header('Content-Type: application/vnd.ms-excel'),
            'ods' => header('Content-Type: application/vnd.oasis.opendocument.spreadsheet'),
            default => throw new ApplicationException('Unsupported file format'),
        };

        header('Content-Disposition: attachment;filename="'.$options['fileName'].'"');

        $writer->save('php://output');
        exit();
    }
}
