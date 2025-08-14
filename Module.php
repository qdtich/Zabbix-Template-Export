<?php

declare(strict_types = 1);

namespace Modules\TemplateExport;

use APP, CController, CWebUser, CMenuItem, Zabbix\Core\CModule;

class Module extends CModule {
	public function init(): void {
		$menu = _('Data collection');

		APP::Component()->get('menu.main')->findOrAdd($menu)->getSubmenu()->insertAfter('Templates', (new CMenuItem(_('Template export')))->setAction('template.export'));
	}

	public function onBeforeAction(CController $action): void {}

	public function onTerminate(CController $action): void {}
}