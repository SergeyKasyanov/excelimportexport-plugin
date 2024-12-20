# ExcelImportExport

Plugin for OctoberCMS 3.x

Provides drop-in replacement for ImportExportController/ImportModel/ExportModel with support for XLSX, XLS and ODS files.

## Installation

```shell
composer require skasianov/excelimportexport-plugin
```

## Usage

1. Replace `\Backend\Behaviors\ImportExportController` in your controller with `\SKasianov\ExcelImportExport\Behaviors\ImportExportController`
2. Extend your import models with `\SKasianov\ExcelImportExport\Models\ImportModel` instead of `\Backend\Models\ImportModel`
3. Extend your export models with `\SKasianov\ExcelImportExport\Models\ExportModel` instead of `\Backend\Models\ExportModel`
4. Use OctoberCMS's import/export functionality as usual.
