<?php

namespace SKasianov\ExcelImportExport\Behaviors\ImportExportController;

use ApplicationException;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\IReader;

trait ImportsData
{
    protected function getImportFileColumns()
    {
        if (! $path = $this->getImportFilePath()) {
            return null;
        }

        if (! $fileFormat = post('file_format', 'json')) {
            return null;
        }

        return match ($fileFormat) {
            'json' => $this->getImportFileColumnsFromJson($path),
            'csv', 'csv_custom' => $this->getImportFileColumnsFromCsv($path),
            'xlsx', 'xls', 'ods' => $this->getImportFileColumnsFromSpreadsheet($path),
            default => throw new ApplicationException('Unsupported file format'),
        };
    }

    protected function getImportFileColumnsFromSpreadsheet(string $path)
    {
        $spreadsheet = IOFactory::load($path, IReader::READ_DATA_ONLY, [
            IOFactory::READER_XLSX,
            IOFactory::READER_XLS,
            IOFactory::READER_ODS,
        ]);

        $sheet = $spreadsheet->getActiveSheet();
        $data = $sheet->toArray();

        return $data[0];
    }
}
