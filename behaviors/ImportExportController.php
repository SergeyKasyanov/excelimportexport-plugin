<?php

namespace SKasianov\ExcelImportExport\Behaviors;

use SKasianov\ExcelImportExport\Behaviors\ImportExportController\ExportsData;
use SKasianov\ExcelImportExport\Behaviors\ImportExportController\ImportsData;

class ImportExportController extends \Backend\Behaviors\ImportExportController
{
    use ExportsData;
    use ImportsData;

    public function guessViewPath($suffix = '', $isPublic = false): ?string
    {
        return $this->guessViewPathFrom(parent::class, $suffix, $isPublic);
    }
}
