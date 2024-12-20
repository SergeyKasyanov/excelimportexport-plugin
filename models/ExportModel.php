<?php

namespace SKasianov\ExcelImportExport\Models;

use ApplicationException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use SystemException;

abstract class ExportModel extends \Backend\Models\ExportModel
{
    /**
     * @throws ApplicationException
     * @throws SystemException
     */
    protected function processExportData($columns, $results, $options): string
    {
        // Validate
        if (! $results) {
            throw new ApplicationException(__('There was no data supplied to export'));
        }

        // Extend columns
        $columns = $this->exportExtendColumns($columns);

        // Save for download
        /** @noinspection NonSecureUniqidUsageInspection */
        $fileName = uniqid('oc');

        $fileName .= match ($this->file_format) {
            'json' => 'xjson',
            'csv', 'csv_custom' => 'xcsv',
            'xlsx' => 'xxlsx',
            'xls' => 'xxls',
            'ods' => 'xods',
            default => throw new ApplicationException('Unsupported file format'),
        };

        $options['savePath'] = $this->getTemporaryExportPath($fileName);

        switch ($this->file_format) {
            case 'json':
                $this->processExportDataAsJson($columns, $results, $options);
                break;
            case 'csv':
            case 'csv_custom':
                $this->processExportDataAsCsv($columns, $results, $options);
                break;
            case 'xlsx':
            case 'xls':
            case 'ods':
                $this->processExportDataAsSpreadsheet($columns, $results, $options);
                break;
            default:
                throw new ApplicationException('Unsupported file format');
        }

        return $fileName;
    }

    /**
     * @throws ApplicationException
     */
    protected function processExportDataAsSpreadsheet($columns, $results, $options): Ods|Xlsx|Xls
    {
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
                $sheet->setCellValue("$letter$rowNum", $column);
                $colNum++;
            }

            $rowNum = 2;
        }

        foreach ($results as $record) {
            $colNum = 0;
            $rowData = $this->matchDataToColumns($record, $columns);
            foreach ($rowData as $value) {
                $letter = $alphabet[$colNum];
                $sheet->setCellValue("$letter$rowNum", $value);
                $colNum++;
            }
            $rowNum++;
        }

        $xlsx = match ($options['fileFormat']) {
            'xlsx' => new Xlsx($spreadsheet),
            'xls' => new Xls($spreadsheet),
            'ods' => new Ods($spreadsheet),
            default => throw new ApplicationException('Unsupported file format'),
        };

        if ($options['useOutput']) {
            match ($options['fileFormat']) {
                'xlsx' => header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
                'xls' => header('Content-Type: application/vnd.ms-excel'),
                'ods' => header('Content-Type: application/vnd.oasis.opendocument.spreadsheet'),
                default => throw new ApplicationException('Unsupported file format'),
            };

            header('Content-Disposition: attachment;filename="'.$options['fileName'].'"');

            $xlsx->save('php://output');
        } elseif ($path = $options['savePath']) {
            $xlsx->save($path);
        }

        return $xlsx;
    }
}
