<?php

$html_page = (new CHtmlPage())->setTitle(_('Template export'));

$form_list = (new CFormList())
    ->addRow(
        (new CDiv(_('This form allows you to export templates using Excel file.')))
    )
    ->addRow(
        (new CLabel(_('Template for export'), 'export_templateid'))->setAsteriskMark(),
        (new CMultiSelect([
            'name' => 'export_templateid',
            'object_name' => 'templates',
            'data' => '',
            'multiple' => false,
            'popup' => [
                'parameters' => [
                    'srctbl' => 'templates',
                    'srcfld1' => 'hostid',
                    'srcfld2' => 'host',
                    'dstfrm' => 'templateExportForm',
                    'dstfld1' => 'export_templateid',
                    'normal_only' => '1'
                ]
            ]
        ]))->setWidth(ZBX_TEXTAREA_MEDIUM_WIDTH)
    );

$form = (new CForm())
    ->setId('template-export-form')
    ->setName('templateExportForm')
    ->setAction((new CUrl('zabbix.php'))
        ->setArgument('action', 'template.output')
        ->getUrl()
    )
    ->addItem(
        (new CTabView())
            ->addTab('template.export', _('Template Export'), $form_list)
            ->setFooter(makeFormFooter(
                new CSubmit('export', _('Export')),
                [(new CSimpleButton(_('Reset')))->onClick("document.location = " . json_encode((new CUrl('zabbix.php'))->setArgument('action', 'template.export')->getUrl()))]
            ))
    );

$html_page->addItem($form)->show();
