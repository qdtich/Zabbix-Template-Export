<?php

declare(strict_types = 0);

namespace Modules\TemplateExport\Actions;

use CController, CControllerResponseData, CRoleHelper;

class TemplateExport extends CController {
    public function init(): void {
        $this->disableCsrfValidation();
    }

    protected function checkInput(): bool {
        return true;
    }

    protected function checkPermissions(): bool {
        return $this->checkAccess(CRoleHelper::UI_ADMINISTRATION_GENERAL);
    }

    protected function doAction(): void {
        $response = new CControllerResponseData([]);
        $this->setResponse($response);
    }
}