<?php

declare(strict_types = 0);

namespace Modules\TemplateExport\Actions;

use CController, CControllerResponseData, CRoleHelper, API, CArrayHelper;

class TemplateOutput extends CController {
    public function init(): void {
        $this->disableCsrfValidation();
    }

    protected function checkInput(): bool {
		$fields = [
			'export_templateid' => 'db hosts.hostid'
		];

		$ret = $this->validateInput($fields);

		if (!$ret) {
			$this->setResponse(new CControllerResponseFatal());
		}

		return $ret;
	}

    protected function checkPermissions(): bool {
		return $this->checkAccess(CRoleHelper::UI_ADMINISTRATION_GENERAL);
	}

    protected function doAction(): void {
		$data = ['export_templateid' => $this->getInput('export_templateid')];

		$data['export_template_data'] = API::Template()->get([
			'output' => ['templateid', 'name'],
			'templateids' => $data['export_templateid']
		]);

		$data['export_template_data'] = CArrayHelper::renameObjectsKeys($data['export_template_data'], ['templateid' => 'id']);

		$data['export_item_data'] = API::Item()->get([
			'output' => ['itemid', 'hostid', 'name', 'key_', 'type', 'delay', 'history', 'trends', 'status', 'state', 'value_type', 'units', 'snmp_oid', 'description', 'master_itemid'],
			'selectPreprocessing' => 'extend',
			'templateids' => $data['export_template_data'][0]['id'],
			'sortfield' => 'name'
		]);

		$data['export_trigger_data'] = API::Trigger()->get([
			'output' => ['triggerid', 'expression', 'description', 'flags', 'type', 'status', 'state', 'value', 'priority', 'recovery_mode', 'recovery_expression', 'correlation_mode', 'manual_close', 'opdata', 'event_name', 'comments'],
			'templateids' => $data['export_template_data'][0]['id'],
			'selectDependencies' => 'extend',
			'sortfield' => 'description'
		]);

		$data['export_macro_data'] = API::UserMacro()->get([
			'output' => ['macro', 'value', 'type', 'automatic', 'description'],
			'hostids' => $data['export_template_data'][0]['id'],
			'sortfield' => 'macro'
		]);

		$data['export_drule_data'] = API::DiscoveryRule()->get([
			'output' => ['itemid', 'type', 'snmp_oid', 'hostid', 'name', 'key_', 'delay', 'status', 'state', 'description', 'master_itemid', 'lifetime', 'lifetime_type', 'enabled_lifetime', 'enabled_lifetime_type'],
			'templateids' => $data['export_template_data'][0]['id'],
			'sortfield' => 'name'
		]);

		$data['export_item_proto_data'] = [];
		$data['export_trigger_proto_data'] = [];
		$data['export_host_proto_data'] = [];
		foreach ($data['export_drule_data'] as $drules) {
			$ip_tmp = API::ItemPrototype()->get([
				'output' => ['itemid', 'name', 'key_', 'type', 'delay', 'history', 'trends', 'status', 'value_type', 'units', 'snmp_oid', 'description', 'master_itemid'],
				'selectPreprocessing' => 'extend',
				'discoveryids' => $drules['itemid'],
				'sortfield' => 'name'
			]);
			if ($ip_tmp != []) {
				array_push($data['export_item_proto_data'], $ip_tmp);
			}

			$tp_tmp = API::TriggerPrototype()->get([
				'output' => ['triggerid', 'expression', 'description', 'flags', 'type', 'status', 'state', 'value', 'priority', 'comments', 'recovery_mode', 'recovery_expression', 'correlation_mode', 'manual_close', 'opdata', 'event_name'],
				'discoveryids' => $drules['itemid'],
				'selectDependencies' => 'extend',
				'sortfield' => 'description'
			]);
			if ($tp_tmp != []) {
				array_push($data['export_trigger_proto_data'], $tp_tmp);
			}

			$hp_tmp = API::HostPrototype()->get([
				'output' => 'extend',
				'discoveryids' => $drules['itemid'],
				'selectGroupLinks' => 'extend',
				'sortfield' => 'name'
			]);
			if ($hp_tmp != []) {
				array_push($data['export_host_proto_data'], $hp_tmp);
			}
		}

		$response = new CControllerResponseData($data);
		$this->setResponse($response);
	}
}