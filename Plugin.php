<?php

namespace SKasianov\ExcelImportExport;

use Event;
use Request;
use System\Classes\PluginBase;

class Plugin extends PluginBase
{
    public function pluginDetails(): array
    {
        return [
            'name' => 'ExcelImportExport',
            'description' => 'Provides replacement for ImportExportController with XLSX/XLS/ODS support.',
            'author' => 'SKasianov',
            'icon' => 'icon-file-excel',
        ];
    }

    public function register(): void
    {
        $module = Request::segment(2);
        if ($module === 'tailor') {
            return;
        }

        Event::listen('system.extendConfigFile', function (string $path, array $config) {
            if ($path === '/modules/backend/behaviors/importexportcontroller/partials/fields_export.yaml') {
                $config['fields']['file_format']['options']['xlsx'] = 'XLSX (MS Office 2007 and above)';
                $config['fields']['file_format']['options']['xls'] = 'XLS (MS Office 95/97/2003)';
                $config['fields']['file_format']['options']['ods'] = 'ODS (OpenOffice/LibreOffice)';
            }

            if ($path === '/modules/backend/behaviors/importexportcontroller/partials/fields_import.yaml') {
                $config['fields']['file_format']['options']['xlsx'] = 'XLSX/XLS/ODS (MS Office/OpenOffice/LibreOffice)';

                $config['fields']['import_file']['fileTypes'] = ['csv', 'json', 'xlsx', 'xls', 'ods'];

                $config['fields']['first_row_titles']['trigger']['condition'] = 'value[csv][csv_custom][xlsx]';
            }

            return $config;
        });
    }
}
