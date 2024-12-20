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
            'description' => 'No description provided yet...',
            'author' => 'SKasianov',
            'icon' => 'icon-leaf',
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
            }

            return $config;
        });
    }
}
