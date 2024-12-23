<?php

namespace SKasianov\ExcelImportExport\Models;

use ApplicationException;
use Arr;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\IReader;

abstract class ImportModel extends \Backend\Models\ImportModel
{
    protected function processImportData($filePath, $matches, $options)
    {
        return match ($this->file_format) {
            'json' => $this->processImportDataAsJson($filePath, $matches, $options),
            'csv', 'csv_custom' => $this->processImportDataAsCsv($filePath, $matches, $options),
            'xlsx', 'xls', 'ods' => $this->processImportDataFromSpreadsheet($filePath, $matches, $options),
            default => throw new ApplicationException('Unsupported file format'),
        };
    }

    protected function processImportDataFromSpreadsheet($filePath, $matches, $options)
    {
        $spreadsheet = IOFactory::load($filePath, IReader::READ_DATA_ONLY, [
            IOFactory::READER_XLSX,
            IOFactory::READER_XLS,
            IOFactory::READER_ODS,
        ]);

        $sheet = $spreadsheet->getActiveSheet();

        $sheetData = $sheet->toArray();
        if ($options['firstRowTitles']) {
            $sheetData = array_slice($sheetData, 1);
        }

        $result = [];
        foreach ($sheetData as $rowData) {
            $result[] = $this->processSpreadsheetImportRow($rowData, $matches);
        }

        return $result;
    }

    protected function processSpreadsheetImportRow($rowData, $matches): array
    {
        $newRow = [];

        foreach ($matches as $columnIndex => $dbNames) {
            $value = Arr::get($rowData, $columnIndex);
            foreach ((array) $dbNames as $dbName) {
                $newRow[$dbName] = $value;
            }
        }

        return $newRow;
    }
}
