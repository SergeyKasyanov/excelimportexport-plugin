<?php

namespace SKasianov\ExcelImportExport\Behaviors\ImportExportController;

use ApplicationException;
use Str;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\IReader;

trait ImportsData
{
    public function actionImportLoadColumnSampleForm()
    {
        if (($columnId = post('file_column_id', false)) === false) {
            throw new ApplicationException(__("Missing column identifier"));
        }

        $columns = $this->getImportFileColumns();
        if (!array_key_exists($columnId, $columns)) {
            throw new ApplicationException(__("Unknown column"));
        }

        $path = $this->getImportFilePath();

        if (!$fileFormat = post('file_format', 'json')) {
            return null;
        }

        $data = match ($fileFormat) {
            'json' => $this->getImportSampleColumnsFromJson($path, (int) $columnId),
            'csv', 'csv_custom' => $this->getImportSampleColumnsFromCsv($path, (int) $columnId),
            'xlsx', 'xls', 'ods' => $this->getImportSampleColumnsFromSpreadsheet($path, (int) $columnId),
            default => throw new ApplicationException('Unsupported file format'),
        };

        // Clean up data
        foreach ($data as $index => $sample) {
            $data[$index] = Str::limit($sample, 100);
            if (!strlen($data[$index])) {
                unset($data[$index]);
            }
        }

        $this->vars['columnName'] = array_get($columns, $columnId);
        $this->vars['columnData'] = $data;
    }

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
        $highestColumn = $sheet->getHighestColumn();

        if (! post('first_row_titles')) {
            $headers = [];
            $columnsCount = Coordinate::columnIndexFromString($highestColumn);

            for ($i = 1; $i <= $columnsCount; $i++) {
                $headers[] = 'Column #'.$i;
            }

            return $headers;
        }

        $data = $sheet->rangeToArray('A1:'.$highestColumn. 1);

        return $data[0];
    }

    protected function getImportSampleColumnsFromSpreadsheet($path, $columnIndex)
    {
        $spreadsheet = IOFactory::load($path, IReader::READ_DATA_ONLY, [
            IOFactory::READER_XLSX,
            IOFactory::READER_XLS,
            IOFactory::READER_ODS,
        ]);

        // init
        $fromRow = 1;
        $limit = 50;
        $columnIndex++;
        $sheet = $spreadsheet->getActiveSheet();

        // if selected column without any data
        $highestColumn = $sheet->getHighestColumn();
        $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);
        if ($columnIndex > $highestColumnIndex) {
            return [];
        }

        // init range
        if (!post('first_row_titles')) {
            $fromRow++;
        }
        $toRow = $fromRow + $limit;

        // read sample data from selected column
        $selectedColumn = Coordinate::stringFromColumnIndex($columnIndex);
        $data = $sheet->rangeToArray($selectedColumn . $fromRow . ':' . $selectedColumn . $toRow);

        return array_column($data, 0);
    }
}
