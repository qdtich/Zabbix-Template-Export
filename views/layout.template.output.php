<?php

include_once("xlsxwriter.class.php");

$writer = new XLSXWriter();
header('Content-disposition:attachment; filename="' . XLSXWriter::sanitize_filename($data['export_template_data'][0]['name']) . '.xlsx"');
header("Content-Type:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
header('Content-Transfer-Encoding:binary');
header('Cache-Control:must-revalidate');
header('Pragma:public');

$header_styles = array('font'=>'Arial','font-size'=>10,'font-style'=>'bold','fill'=>'#eee','halign'=>'center','valign'=>'center','border'=>'left,right,top,bottom','height'=>50,'wrap_text'=>true,'freeze_rows'=>1,'freeze_columns'=>1);
$row_styles = array('font'=>'Arial','font-size'=>10,'halign'=>'left','valign'=>'center','border'=>'left,right,top,bottom','wrap_text'=>true);

$item_header = array(
  'no.'=>'@',
  'name'=>'@',
  'key_'=>'@',
  'snmp_oid'=>'@',
  'type'=>'@',
  'delay'=>'@',
  'history'=>'@',
  'trends'=>'@',
  'status'=>'@',
  'state'=>'@',
  'value_type'=>'@',
  'units'=>'@',
  'master_item'=>'@',
  'preprocessing_type'=>'@',
  'preprocessing_params'=>'@',
  'description'=>'@'
);

$sheet1 = 'item';

$writer->writeSheetHeader($sheet1, $item_header, array_merge($header_styles, ['widths'=>[5,30,40,30,15,20,8,8,8,8,17,8,15,20, 60, 60]]));

$i = 0;
$a = 0;
foreach ($data['export_item_data'] as $items) {
    $r_name = $items['name'];

    $r_key = $items['key_'];

    if ($items['type'] == 0) {
        $r_type = 'Zabbix agent';
    }
    elseif ($items['type'] == 2) {
        $r_type = 'Zabbix trapper';
    }
    elseif ($items['type'] == 3) {
        $r_type = 'Simple check';
    }
    elseif ($items['type'] == 5) {
        $r_type = 'Zabbix internal';
    }
    elseif ($items['type'] == 7) {
        $r_type = 'Zabbix agent (active)';
    }
    elseif ($items['type'] == 9) {
        $r_type = 'Web item';
    }
    elseif ($items['type'] == 10) {
        $r_type = 'External check';
    }
    elseif ($items['type'] == 11) {
        $r_type = 'Database monitor';
    }
    elseif ($items['type'] == 12) {
        $r_type = 'IPMI agent';
    }
    elseif ($items['type'] == 13) {
        $r_type = 'SSH agent';
    }
    elseif ($items['type'] == 14) {
        $r_type = 'TELNET agent';
    }
    elseif ($items['type'] == 15) {
        $r_type = 'Calculated';
    }
    elseif ($items['type'] == 16) {
        $r_type = 'JMX agent';
    }
    elseif ($items['type'] == 17) {
        $r_type = 'SNMP trap';
    }
    elseif ($items['type'] == 18) {
        $r_type = 'Dependent item';
    }
    elseif ($items['type'] == 19) {
        $r_type = 'HTTP agent';
    }
    elseif ($items['type'] == 20) {
        $r_type = 'SNMP agent';
    }
    elseif ($items['type'] == 21) {
        $r_type = 'Script';
    }
    elseif ($items['type'] == 22) {
        $r_type = 'Browser';
    }

    $r_delay = $items['delay'];

    $r_history = $items['history'];

    $r_trends = $items['trends'];

    if ($items['status'] == 0) {
        $r_status = 'enabled';
    }
    elseif ($items['status'] == 1) {
        $r_status = 'disabled';
    }
    if ($items['state'] == 0) {
        $r_state = 'normal';
    }
    elseif ($items['state'] == 1) {
        $r_state = 'not supported';
    }
    if ($items['value_type'] == 0) {
        $r_value_type = 'numeric float';
    }
    elseif ($items['value_type'] == 1) {
        $r_value_type = 'character';
    }
    elseif ($items['value_type'] == 2) {
        $r_value_type = 'log';
    }
    elseif ($items['value_type'] == 3) {
        $r_value_type = 'numeric unsigned';
    }
    elseif ($items['value_type'] == 4) {
        $r_value_type = 'text';
    }
    elseif ($items['value_type'] == 5) {
        $r_value_type = 'binary';
    }

    $r_units = $items['units'];

    $r_snmp_oid = $items['snmp_oid'];

    $db_mi_res = DBfetch(DBselect('select name from items where itemid=' . $items['master_itemid']));
    $r_master_item = $db_mi_res['name'];

    $r_description = $items['description'];
    
    $num = count($items['preprocessing']);
    if ($num > 1) {
        $i++;
        for ($x=0; $x<$num; $x++) {
            if ( $items['preprocessing'][$x]['type'] == 1) {
                $r_pp_type = 'Custom multiplier';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 2) {
                $r_pp_type = 'Right trim';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 3) {
                $r_pp_type = 'Left trim';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 4) {
                $r_pp_type = 'Trim';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 5) {
                $r_pp_type = 'Regular expression';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 6) {
                $r_pp_type = 'Boolean to decimal';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 7) {
                $r_pp_type = 'Octal to decimal';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 8) {
                $r_pp_type = 'Hexadecimal to decimal';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 9) {
                $r_pp_type = 'Simple change';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 10) {
                $r_pp_type = 'Change per second';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 11) {
                $r_pp_type = 'XML XPath';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 12) {
                $r_pp_type = 'JSONPath';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 13) {
                $r_pp_type = 'In range';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 14) {
                $r_pp_type = 'Matches regular expression';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 15) {
                $r_pp_type = 'Does not match regular expression';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 16) {
                $r_pp_type = 'Check for error in JSON';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 17) {
                $r_pp_type = 'Check for error in XML';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 18) {
                $r_pp_type = 'Check for error using regular expression';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 19) {
                $r_pp_type = 'Discard unchanged';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 20) {
                $r_pp_type = 'Discard unchanged with heartbeat';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 21) {
                $r_pp_type = 'JavaScript';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 22) {
                $r_pp_type = 'Prometheus pattern';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 23) {
                $r_pp_type = 'Prometheus to JSON';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 24) {
                $r_pp_type = 'CSV to JSON';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 25) {
                $r_pp_type = 'Replace';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 26) {
                $r_pp_type = 'Check unsupported';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 27) {
                $r_pp_type = 'XML to JSON';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 28) {
                $r_pp_type = 'SNMP walk value';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 29) {
                $r_pp_type = 'SNMP walk to JSON';
            }
            elseif ( $items['preprocessing'][$x]['type'] == 30) {
                $r_pp_type = 'SNMP get value';
            }

            $r_pp_params = $items['preprocessing'][$x]['params'];

            if ($x == 0) {
                $writer->writeSheetRow($sheet1, array($i, $r_name, $r_key, $r_snmp_oid, $r_type, $r_delay, $r_history, $r_trends, $r_status, $r_state, $r_value_type, $r_units, $r_master_item, $r_pp_type, $r_pp_params, $r_description), $row_styles);
            }
            else {
                $writer->writeSheetRow($sheet1, array($i, '', '', '', '', '', '', '', '', '', '', '', '', $r_pp_type, $r_pp_params, ''), $row_styles);
            }

            $a++;
        }

        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=0, $end_row=$a, $end_col=0);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=1, $end_row=$a, $end_col=1);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=2, $end_row=$a, $end_col=2);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=3, $end_row=$a, $end_col=3);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=4, $end_row=$a, $end_col=4);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=5, $end_row=$a, $end_col=5);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=6, $end_row=$a, $end_col=6);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=7, $end_row=$a, $end_col=7);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=8, $end_row=$a, $end_col=8);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=9, $end_row=$a, $end_col=9);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=10, $end_row=$a, $end_col=10);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=11, $end_row=$a, $end_col=11);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=12, $end_row=$a, $end_col=12);
        $writer->markMergedCell($sheet1, $start_row=$a-($num-1), $start_col=15, $end_row=$a, $end_col=15);
    }
    elseif ($num == 1) {
        $i++;
        $a++;
        if ($items['preprocessing'][0]['type'] == 1) {
            $r_pp_type = 'Custom multiplier';
        }
        elseif ($items['preprocessing'][0]['type'] == 2) {
            $r_pp_type = 'Right trim';
        }
        elseif ($items['preprocessing'][0]['type'] == 3) {
            $r_pp_type = 'Left trim';
        }
        elseif ($items['preprocessing'][0]['type'] == 4) {
            $r_pp_type = 'Trim';
        }
        elseif ($items['preprocessing'][0]['type'] == 5) {
            $r_pp_type = 'Regular expression';
        }
        elseif ($items['preprocessing'][0]['type'] == 6) {
            $r_pp_type = 'Boolean to decimal';
        }
        elseif ($items['preprocessing'][0]['type'] == 7) {
            $r_pp_type = 'Octal to decimal';
        }
        elseif ($items['preprocessing'][0]['type'] == 8) {
            $r_pp_type = 'Hexadecimal to decimal';
        }
        elseif ($items['preprocessing'][0]['type'] == 9) {
            $r_pp_type = 'Simple change';
        }
        elseif ($items['preprocessing'][0]['type'] == 10) {
            $r_pp_type = 'Change per second';
        }
        elseif ($items['preprocessing'][0]['type'] == 11) {
            $r_pp_type = 'XML XPath';
        }
        elseif ($items['preprocessing'][0]['type'] == 12) {
            $r_pp_type = 'JSONPath';
        }
        elseif ($items['preprocessing'][0]['type'] == 13) {
            $r_pp_type = 'In range';
        }
        elseif ($items['preprocessing'][0]['type'] == 14) {
            $r_pp_type = 'Matches regular expression';
        }
        elseif ($items['preprocessing'][0]['type'] == 15) {
            $r_pp_type = 'Does not match regular expression';
        }
        elseif ($items['preprocessing'][0]['type'] == 16) {
            $r_pp_type = 'Check for error in JSON';
        }
        elseif ($items['preprocessing'][0]['type'] == 17) {
            $r_pp_type = 'Check for error in XML';
        }
        elseif ($items['preprocessing'][0]['type'] == 18) {
            $r_pp_type = 'Check for error using regular expression';
        }
        elseif ($items['preprocessing'][0]['type'] == 19) {
            $r_pp_type = 'Discard unchanged';
        }
        elseif ($items['preprocessing'][0]['type'] == 20) {
            $r_pp_type = 'Discard unchanged with heartbeat';
        }
        elseif ($items['preprocessing'][0]['type'] == 21) {
            $r_pp_type = 'JavaScript';
        }
        elseif ($items['preprocessing'][0]['type'] == 22) {
            $r_pp_type = 'Prometheus pattern';
        }
        elseif ($items['preprocessing'][0]['type'] == 23) {
            $r_pp_type = 'Prometheus to JSON';
        }
        elseif ($items['preprocessing'][0]['type'] == 24) {
            $r_pp_type = 'CSV to JSON';
        }
        elseif ($items['preprocessing'][0]['type'] == 25) {
            $r_pp_type = 'Replace';
        }
        elseif ($items['preprocessing'][0]['type'] == 26) {
            $r_pp_type = 'Check unsupported';
        }
        elseif ($items['preprocessing'][0]['type'] == 27) {
            $r_pp_type = 'XML to JSON';
        }
        elseif ($items['preprocessing'][0]['type'] == 28) {
            $r_pp_type = 'SNMP walk value';
        }
        elseif ($items['preprocessing'][0]['type'] == 29) {
            $r_pp_type = 'SNMP walk to JSON';
        }
        elseif ($items['preprocessing'][0]['type'] == 30) {
            $r_pp_type = 'SNMP get value';
        }

        $r_pp_params = $items['preprocessing'][0]['params'];

        $writer->writeSheetRow($sheet1, array($i, $r_name, $r_key, $r_snmp_oid, $r_type, $r_delay, $r_history, $r_trends, $r_status, $r_state, $r_value_type, $r_units, $r_master_item, $r_pp_type, $r_pp_params, $r_description), $row_styles);
    }
    else {
        $i++;
        $a++;
        $writer->writeSheetRow($sheet1, array($i, $r_name, $r_key, $r_snmp_oid, $r_type, $r_delay, $r_history, $r_trends, $r_status, $r_state, $r_value_type, $r_units, $r_master_item, '', '', $r_description), $row_styles);
    }
}

$trigger_header = array(
  'no.'=>'@',
  'expression'=>'@',
  'flags'=>'@',
  'type'=>'@',
  'status'=>'@',
  'state'=>'@',
  'value'=>'@',
  'priority'=>'@',
  'recovery_mode'=>'@',
  'recovery_expression'=>'@',
  'correlation_mode'=>'@',
  'manual_close'=>'@',
  'opdata'=>'@',
  'dependencies'=>'@',
  'event_name'=>'@',
  'comments'=>'@'
);

$sheet2 = 'trigger';

$writer->writeSheetHeader($sheet2, $trigger_header, array_merge($header_styles, ['widths'=>[5,60,13,15,8,15,8,11,11,40,12,8,30,30,30,30]]));

$j = 0;
$b = 0;
foreach ($data['export_trigger_data'] as $triggers) {
    $r_expression = '';
    $exp_res_s = explode('{', $triggers['expression']);
    foreach ($exp_res_s as $exp_res) {
        if ($exp_res != '') {
            $db_func_res = DBfetch(DBselect('select itemid,name from functions where functionid=' . explode('}', $exp_res)[0]));
            $db_item_res = DBfetch(DBselect('select key_ from items where itemid=' . $db_func_res['itemid']));
            $r_expression = $r_expression . $db_func_res['name'] . '(/' . $data['export_template_data'][0]['name'] . '/' . $db_item_res['key_'] . ')' . explode('}', $exp_res)[1];
        }
    }

    if ($triggers['flags'] == 0) {
        $r_flags = 'a plain trigger';
    }
    elseif ($triggers['flags'] == 4) {
        $r_flags = 'a discovered trigger';
    }
    
    if ($triggers['type'] == 0) {
        $r_type = 'do not generate multiple events';
    }
    elseif ($triggers['type'] == 1) {
        $r_type = 'generate multiple events';
    }
    
    if ($triggers['status'] == 0) {
        $r_status = 'enabled';
    }
    elseif ($triggers['status'] == 1) {
        $r_status = 'disabled';
    }
    
    if ($triggers['state'] == 0) {
        $r_state = 'trigger state is up to date';
    }
    elseif ($triggers['state'] == 1) {
        $r_state = 'current trigger state is unknown';
    }
    
    if ($triggers['value'] == 0) {
        $r_value = 'OK';
    }
    elseif ($triggers['value'] == 1) {
        $r_value = 'problem';
    }

    if ($triggers['priority'] == 0) {
        $r_priority = 'not classified';
    }
    elseif ($triggers['priority'] == 1) {
        $r_priority = 'information';
    }
    elseif ($triggers['priority'] == 2) {
        $r_priority = 'warning';
    }
    elseif ($triggers['priority'] == 3) {
        $r_priority = 'average';
    }
    elseif ($triggers['priority'] == 4) {
        $r_priority = 'high';
    }
    elseif ($triggers['priority'] == 5) {
        $r_priority = 'disaster';
    }

    if ($triggers['recovery_mode'] == 0) {
        $r_recovery_mode = 'Expression';
    }
    elseif ($triggers['recovery_mode'] == 1) {
        $r_recovery_mode = 'Recovery expression';
    }
    elseif ($triggers['recovery_mode'] == 2) {
        $r_recovery_mode = 'None';
    }

    $r_recovery_expression = '';
    $exp_res_s = explode('{', $triggers['recovery_expression']);
    foreach ($exp_res_s as $exp_res) {
        if ($exp_res != '') {
            $db_func_res = DBfetch(DBselect('select itemid,name from functions where functionid=' . explode('}', $exp_res)[0]));
            $db_item_res = DBfetch(DBselect('select key_ from items where itemid=' . $db_func_res['itemid']));
            $r_recovery_expression = $r_recovery_expression . $db_func_res['name'] . '(/' . $data['export_template_data'][0]['name'] . '/' . $db_item_res['key_'] . ')' . explode('}', $exp_res)[1];
        }
    }

    if ($triggers['correlation_mode'] == 0) {
        $r_correlation_mode = 'All problems';
    }
    elseif ($triggers['correlation_mode'] == 1) {
        $r_correlation_mode = 'All problems if tag values match';
    }

    if ($triggers['manual_close'] == 0) {
        $r_manual_close = 'No';
    }
    elseif ($triggers['manual_close'] == 1) {
        $r_manual_close = 'Yes';
    }

    $r_opdata = $triggers['opdata'];

    $r_event_name = $triggers['event_name'];

    $r_comments = $triggers['comments'];

    $num = count($triggers['dependencies']);
    if ($num > 1) {
        $j++;
        for ($x=0; $x<$num; $x++) {
            $r_trigger_depen_descr = $triggers['dependencies'][$x]['description'];

            if ($x == 0) {
                $writer->writeSheetRow($sheet2, array($j, $r_expression, $r_flags, $r_type, $r_status, $r_state, $r_value, $r_priority, $r_recovery_mode, $r_recovery_expression, $r_correlation_mode, $r_manual_close, $r_opdata, $r_trigger_depen_descr, $r_event_name, $r_comments), $row_styles);
            }
            else {
                $writer->writeSheetRow($sheet2, array($j, '', '', '', '', '', '', '', '', '', '', '', '', $r_trigger_depen_descr, '', ''), $row_styles);
            }

            $b++;
        }

        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=0, $end_row=$b, $end_col=0);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=1, $end_row=$b, $end_col=1);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=2, $end_row=$b, $end_col=2);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=3, $end_row=$b, $end_col=3);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=4, $end_row=$b, $end_col=4);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=5, $end_row=$b, $end_col=5);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=6, $end_row=$b, $end_col=6);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=7, $end_row=$b, $end_col=7);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=8, $end_row=$b, $end_col=8);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=9, $end_row=$b, $end_col=9);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=10, $end_row=$b, $end_col=10);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=11, $end_row=$b, $end_col=11);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=12, $end_row=$b, $end_col=12);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=14, $end_row=$b, $end_col=14);
        $writer->markMergedCell($sheet2, $start_row=$b-($num-1), $start_col=15, $end_row=$b, $end_col=15);
    }
    elseif ($num == 1) {
        $j++;
        $b++;
        $r_trigger_depen_descr = $triggers['dependencies'][0]['description'];
        $writer->writeSheetRow($sheet2, array($j, $r_expression, $r_flags, $r_type, $r_status, $r_state, $r_value, $r_priority, $r_recovery_mode, $r_recovery_expression, $r_correlation_mode, $r_manual_close, $r_opdata, $r_trigger_depen_descr, $r_event_name, $r_comments), $row_styles);
    }
    else {
        $j++;
        $b++;
        $writer->writeSheetRow($sheet2, array($j, $r_expression, $r_flags, $r_type, $r_status, $r_state, $r_value, $r_priority, $r_recovery_mode, $r_recovery_expression, $r_correlation_mode, $r_manual_close, $r_opdata, '', $r_event_name, $r_comments), $row_styles);
    }
}

$macro_header = array(
  'no.'=>'@',
  'macro'=>'@',
  'value'=>'@',
  'type'=>'@',
  'automatic'=>'@',
  'description'=>'@'
);

$sheet3 = 'macro';

$writer->writeSheetHeader($sheet3, $macro_header, array_merge($header_styles, ['widths'=>[5,50,50,15,30,70]]));

$k = 0;
foreach ($data['export_macro_data'] as $macros) {
    $k++;
    $r_macro = $macros['macro'];

    $r_value = $macros['value'];

    if ($macros['type'] == 0) {
        $r_type = 'Text macro';
    }
    elseif ($macros['type'] == 1) {
        $r_type = 'Secret macro';
    }
    elseif ($macros['type'] == 2) {
        $r_type = 'Vault secret';
    }

    if ($macros['automatic'] == 0) {
        $r_automatic = 'Macro is managed by user';
    }
    elseif ($macros['automatic'] == 1) {
        $r_automatic = 'Macro is managed by discovery rule';
    }

    $r_description = $macros['description'];

    $writer->writeSheetRow($sheet3, array($k, $r_macro, $r_value, $r_type, $r_automatic, $r_description), $row_styles);
}

$drule_header = array(
  'no.'=>'@',
  'name'=>'@',
  'key_'=>'@',
  'snmp_oid'=>'@',
  'type'=>'@',
  'delay'=>'@',
  'status'=>'@',
  'state'=>'@',
  'master_item'=>'@',
  'lifetime'=>'@',
  'lifetime_type'=>'@',
  'enabled_lifetime'=>'@',
  'enabled_lifetime_type'=>'@',
  'description'=>'@'
);

$sheet4 = 'drule';

$writer->writeSheetHeader($sheet4, $drule_header, array_merge($header_styles, ['widths'=>[5,30,40,30,15,8,8,8,15,8,20,10,17,45]]));

$l = 0;
foreach ($data['export_drule_data'] as $drule) {
    $l++;
    if ($drule['type'] == 0) {
        $r_type = 'Zabbix agent';
    }
    elseif ($drule['type'] == 2) {
        $r_type = 'Zabbix trapper';
    }
    elseif ($drule['type'] == 3) {
        $r_type = 'Simple check';
    }
    elseif ($drule['type'] == 5) {
        $r_type = 'Zabbix internal';
    }
    elseif ($drule['type'] == 7) {
        $r_type = 'Zabbix agent (active)';
    }
    elseif ($drule['type'] == 9) {
        $r_type = 'Web item';
    }
    elseif ($drule['type'] == 10) {
        $r_type = 'External check';
    }
    elseif ($drule['type'] == 11) {
        $r_type = 'Database monitor';
    }
    elseif ($drule['type'] == 12) {
        $r_type = 'IPMI agent';
    }
    elseif ($drule['type'] == 13) {
        $r_type = 'SSH agent';
    }
    elseif ($drule['type'] == 14) {
        $r_type = 'TELNET agent';
    }
    elseif ($drule['type'] == 15) {
        $r_type = 'Calculated';
    }
    elseif ($drule['type'] == 16) {
        $r_type = 'JMX agent';
    }
    elseif ($drule['type'] == 17) {
        $r_type = 'SNMP trap';
    }
    elseif ($drule['type'] == 18) {
        $r_type = 'Dependent item';
    }
    elseif ($drule['type'] == 19) {
        $r_type = 'HTTP agent';
    }
    elseif ($drule['type'] == 20) {
        $r_type = 'SNMP agent';
    }
    elseif ($drule['type'] == 21) {
        $r_type = 'Script';
    }
    elseif ($drule['type'] == 22) {
        $r_type = 'Browser';
    }

    $r_snmp_oid = $drule['snmp_oid'];

    $r_name = $drule['name'];
    
    $r_key = $drule['key_'];

    $r_delay = $drule['delay'];

    if ($drule['status'] == 0) {
        $r_status = 'enabled';
    }
    elseif ($drule['status'] == 1) {
        $r_status = 'disabled';
    }

    if ($drule['state'] == 0) {
        $r_state = 'normal';
    }
    elseif ($drule['state'] == 1) {
        $r_state = 'not supported';
    }

    $db_mi_res = DBfetch(DBselect('select name from items where itemid=' . $drule['master_itemid']));
    $r_master_item = $db_mi_res['name'];

    $r_lifetime = $drule['lifetime'];

    if ($drule['lifetime_type'] == 0) {
        $r_lifetime_type = 'Delete after lifetime threshold is reached';
    }
    elseif ($drule['lifetime_type'] == 1) {
        $r_lifetime_type = 'Do not delete';
    }
    elseif ($drule['lifetime_type'] == 2) {
        $r_lifetime_type = 'Delete immediately';
    }

    $r_enabled_lifetime = $drule['enabled_lifetime'];

    if ($drule['enabled_lifetime_type'] == 0) {
        $r_enabled_lifetime_type = 'Disable after lifetime threshold is reached';
    }
    elseif ($drule['enabled_lifetime_type'] == 1) {
        $r_enabled_lifetime_type = 'Do not disable';
    }
    elseif ($drule['enabled_lifetime_type'] == 2) {
        $r_enabled_lifetime_type = 'Disable immediately';
    }

    $r_description = $drule['description'];

    $writer->writeSheetRow($sheet4, array($l, $r_name, $r_key, $r_snmp_oid, $r_type, $r_delay, $r_status, $r_state, $r_master_item, $r_lifetime, $r_lifetime_type, $r_enabled_lifetime, $r_enabled_lifetime_type, $r_description), $row_styles);
}

$item_proto_header = array(
  'no.'=>'@',
  'name'=>'@',
  'key_'=>'@',
  'snmp_oid'=>'@',
  'type'=>'@',
  'delay'=>'@',
  'history'=>'@',
  'trends'=>'@',
  'status'=>'@',
  'state'=>'@',
  'value_type'=>'@',
  'units'=>'@',
  'master_item'=>'@',
  'preprocessing_type'=>'@',
  'preprocessing_params'=>'@',
  'description'=>'@'
);

$sheet5 = 'item proto';

$writer->writeSheetHeader($sheet5, $item_header, array_merge($header_styles, ['widths'=>[5,30,40,30,15,20,8,8,8,8,17,8,15,20, 60, 60]]));

$m = 0;
$c = 0;
foreach ($data['export_item_proto_data'] as $items) {
    foreach ($items as $item) {
        $r_name = $item['name'];

        $r_key = $item['key_'];

        if ($item['type'] == 0) {
            $r_type = 'Zabbix agent';
        }
        elseif ($item['type'] == 2) {
            $r_type = 'Zabbix trapper';
        }
        elseif ($item['type'] == 3) {
            $r_type = 'Simple check';
        }
        elseif ($item['type'] == 5) {
            $r_type = 'Zabbix internal';
        }
        elseif ($item['type'] == 7) {
            $r_type = 'Zabbix agent (active)';
        }
        elseif ($item['type'] == 9) {
            $r_type = 'Web item';
        }
        elseif ($item['type'] == 10) {
            $r_type = 'External check';
        }
        elseif ($item['type'] == 11) {
            $r_type = 'Database monitor';
        }
        elseif ($item['type'] == 12) {
            $r_type = 'IPMI agent';
        }
        elseif ($item['type'] == 13) {
            $r_type = 'SSH agent';
        }
        elseif ($item['type'] == 14) {
            $r_type = 'TELNET agent';
        }
        elseif ($item['type'] == 15) {
            $r_type = 'Calculated';
        }
        elseif ($item['type'] == 16) {
            $r_type = 'JMX agent';
        }
        elseif ($item['type'] == 17) {
            $r_type = 'SNMP trap';
        }
        elseif ($item['type'] == 18) {
            $r_type = 'Dependent item';
        }
        elseif ($item['type'] == 19) {
            $r_type = 'HTTP agent';
        }
        elseif ($item['type'] == 20) {
            $r_type = 'SNMP agent';
        }
        elseif ($item['type'] == 21) {
            $r_type = 'Script';
        }
        elseif ($item['type'] == 22) {
            $r_type = 'Browser';
        }

        $r_delay = $item['delay'];

        $r_history = $item['history'];

        $r_trends = $item['trends'];

        if ($item['status'] == 0) {
            $r_status = 'enabled';
        }
        elseif ($item['status'] == 1) {
            $r_status = 'disabled';
        }

        if ($item['state'] == 0) {
            $r_state = 'normal';
        }
        elseif ($item['state'] == 1) {
            $r_state = 'not supported';
        }

        if ($item['value_type'] == 0) {
            $r_value_type = 'numeric float';
        }
        elseif ($item['value_type'] == 1) {
            $r_value_type = 'character';
        }
        elseif ($item['value_type'] == 2) {
            $r_value_type = 'log';
        }
        elseif ($item['value_type'] == 3) {
            $r_value_type = 'numeric unsigned';
        }
        elseif ($item['value_type'] == 4) {
            $r_value_type = 'text';
        }
        elseif ($item['value_type'] == 5) {
            $r_value_type = 'binary';
        }

        $r_units = $item['units'];

        $r_snmp_oid = $item['snmp_oid'];

        $db_mi_res = DBfetch(DBselect('select name from items where itemid=' . $item['master_itemid']));
        $r_master_item = $db_mi_res['name'];

        $r_description = $item['description'];
        
        $num = count($item['preprocessing']);
        if ($num > 1) {
            $m++;
            for ($x=0; $x<$num; $x++) {
                if ( $item['preprocessing'][$x]['type'] == 1) {
                    $r_pp_type = 'Custom multiplier';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 2) {
                    $r_pp_type = 'Right trim';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 3) {
                    $r_pp_type = 'Left trim';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 4) {
                    $r_pp_type = 'Trim';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 5) {
                    $r_pp_type = 'Regular expression';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 6) {
                    $r_pp_type = 'Boolean to decimal';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 7) {
                    $r_pp_type = 'Octal to decimal';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 8) {
                    $r_pp_type = 'Hexadecimal to decimal';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 9) {
                    $r_pp_type = 'Simple change';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 10) {
                    $r_pp_type = 'Change per second';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 11) {
                    $r_pp_type = 'XML XPath';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 12) {
                    $r_pp_type = 'JSONPath';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 13) {
                    $r_pp_type = 'In range';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 14) {
                    $r_pp_type = 'Matches regular expression';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 15) {
                    $r_pp_type = 'Does not match regular expression';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 16) {
                    $r_pp_type = 'Check for error in JSON';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 17) {
                    $r_pp_type = 'Check for error in XML';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 18) {
                    $r_pp_type = 'Check for error using regular expression';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 19) {
                    $r_pp_type = 'Discard unchanged';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 20) {
                    $r_pp_type = 'Discard unchanged with heartbeat';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 21) {
                    $r_pp_type = 'JavaScript';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 22) {
                    $r_pp_type = 'Prometheus pattern';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 23) {
                    $r_pp_type = 'Prometheus to JSON';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 24) {
                    $r_pp_type = 'CSV to JSON';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 25) {
                    $r_pp_type = 'Replace';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 26) {
                    $r_pp_type = 'Check unsupported';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 27) {
                    $r_pp_type = 'XML to JSON';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 28) {
                    $r_pp_type = 'SNMP walk value';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 29) {
                    $r_pp_type = 'SNMP walk to JSON';
                }
                elseif ( $item['preprocessing'][$x]['type'] == 30) {
                    $r_pp_type = 'SNMP get value';
                }

                $r_pp_params = $item['preprocessing'][$x]['params'];

                $c++;

                if ($x == 0) {
                    $writer->writeSheetRow($sheet5, array($m, $r_name, $r_key, $r_snmp_oid, $r_type, $r_delay, $r_history, $r_trends, $r_status, $r_state, $r_value_type, $r_units, $r_master_item, $r_pp_type, $r_pp_params, $r_description), $row_styles);
                }
                else {
                    $writer->writeSheetRow($sheet5, array($m, '', '', '', '', '', '', '', '', '', '', '', '', $r_pp_type, $r_pp_params, ''), $row_styles);
                }
            }

            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=0, $end_row=$c, $end_col=0);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=1, $end_row=$c, $end_col=1);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=2, $end_row=$c, $end_col=2);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=3, $end_row=$c, $end_col=3);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=4, $end_row=$c, $end_col=4);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=5, $end_row=$c, $end_col=5);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=6, $end_row=$c, $end_col=6);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=7, $end_row=$c, $end_col=7);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=8, $end_row=$c, $end_col=8);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=9, $end_row=$c, $end_col=9);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=10, $end_row=$c, $end_col=10);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=11, $end_row=$c, $end_col=11);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=12, $end_row=$c, $end_col=12);
            $writer->markMergedCell($sheet5, $start_row=$c-($num-1), $start_col=15, $end_row=$c, $end_col=15);
        }
        elseif ($num == 1) {
            $m++;
            $c++;
            if ( $item['preprocessing'][0]['type'] == 1) {
                $r_pp_type = 'Custom multiplier';
            }
            elseif ( $item['preprocessing'][0]['type'] == 2) {
                $r_pp_type = 'Right trim';
            }
            elseif ( $item['preprocessing'][0]['type'] == 3) {
                $r_pp_type = 'Left trim';
            }
            elseif ( $item['preprocessing'][0]['type'] == 4) {
                $r_pp_type = 'Trim';
            }
            elseif ( $item['preprocessing'][0]['type'] == 5) {
                $r_pp_type = 'Regular expression';
            }
            elseif ( $item['preprocessing'][0]['type'] == 6) {
                $r_pp_type = 'Boolean to decimal';
            }
            elseif ( $item['preprocessing'][0]['type'] == 7) {
                $r_pp_type = 'Octal to decimal';
            }
            elseif ( $item['preprocessing'][0]['type'] == 8) {
                $r_pp_type = 'Hexadecimal to decimal';
            }
            elseif ( $item['preprocessing'][0]['type'] == 9) {
                $r_pp_type = 'Simple change';
            }
            elseif ( $item['preprocessing'][0]['type'] == 10) {
                $r_pp_type = 'Change per second';
            }
            elseif ( $item['preprocessing'][0]['type'] == 11) {
                $r_pp_type = 'XML XPath';
            }
            elseif ( $item['preprocessing'][0]['type'] == 12) {
                $r_pp_type = 'JSONPath';
            }
            elseif ( $item['preprocessing'][0]['type'] == 13) {
                $r_pp_type = 'In range';
            }
            elseif ( $item['preprocessing'][0]['type'] == 14) {
                $r_pp_type = 'Matches regular expression';
            }
            elseif ( $item['preprocessing'][0]['type'] == 15) {
                $r_pp_type = 'Does not match regular expression';
            }
            elseif ( $item['preprocessing'][0]['type'] == 16) {
                $r_pp_type = 'Check for error in JSON';
            }
            elseif ( $item['preprocessing'][0]['type'] == 17) {
                $r_pp_type = 'Check for error in XML';
            }
            elseif ( $item['preprocessing'][0]['type'] == 18) {
                $r_pp_type = 'Check for error using regular expression';
            }
            elseif ( $item['preprocessing'][0]['type'] == 19) {
                $r_pp_type = 'Discard unchanged';
            }
            elseif ( $item['preprocessing'][0]['type'] == 20) {
                $r_pp_type = 'Discard unchanged with heartbeat';
            }
            elseif ( $item['preprocessing'][0]['type'] == 21) {
                $r_pp_type = 'JavaScript';
            }
            elseif ( $item['preprocessing'][0]['type'] == 22) {
                $r_pp_type = 'Prometheus pattern';
            }
            elseif ( $item['preprocessing'][0]['type'] == 23) {
                $r_pp_type = 'Prometheus to JSON';
            }
            elseif ( $item['preprocessing'][0]['type'] == 24) {
                $r_pp_type = 'CSV to JSON';
            }
            elseif ( $item['preprocessing'][0]['type'] == 25) {
                $r_pp_type = 'Replace';
            }
            elseif ( $item['preprocessing'][0]['type'] == 26) {
                $r_pp_type = 'Check unsupported';
            }
            elseif ( $item['preprocessing'][0]['type'] == 27) {
                $r_pp_type = 'XML to JSON';
            }
            elseif ( $item['preprocessing'][0]['type'] == 28) {
                $r_pp_type = 'SNMP walk value';
            }
            elseif ( $item['preprocessing'][0]['type'] == 29) {
                $r_pp_type = 'SNMP walk to JSON';
            }
            elseif ( $item['preprocessing'][0]['type'] == 30) {
                $r_pp_type = 'SNMP get value';
            }

            $r_pp_params = $item['preprocessing'][0]['params'];

            $writer->writeSheetRow($sheet5, array($m, $r_name, $r_key, $r_snmp_oid, $r_type, $r_delay, $r_history, $r_trends, $r_status, $r_state, $r_value_type, $r_units, $r_master_item, $r_pp_type, $r_pp_params, $r_description), $row_styles);
        }
        else {
            $m++;
            $c++;
            $writer->writeSheetRow($sheet5, array($m, $r_name, $r_key, $r_snmp_oid, $r_type, $r_delay, $r_history, $r_trends, $r_status, $r_state, $r_value_type, $r_units, $r_master_item, '', '', $r_description), $row_styles);
        }
    }
}

$trigger_proto_header = array(
  'no.'=>'@',
  'expression'=>'@',
  'flags'=>'@',
  'type'=>'@',
  'status'=>'@',
  'state'=>'@',
  'value'=>'@',
  'priority'=>'@',
  'recovery_mode'=>'@',
  'recovery_expression'=>'@',
  'correlation_mode'=>'@',
  'manual_close'=>'@',
  'opdata'=>'@',
  'dependencies'=>'@',
  'event_name'=>'@',
  'comments'=>'@'
);

$sheet6 = 'trigger proto';

$writer->writeSheetHeader($sheet6, $trigger_header, array_merge($header_styles, ['widths'=>[5,60,13,15,8,15,8,11,11,40,12,8,30,30,30,30]]));

$n = 0;
$d = 0;
foreach ($data['export_trigger_proto_data'] as $triggers) {
    foreach ($triggers as $trigger) {
        $r_expression = '';
        $exp_res_s = explode('{', $trigger['expression']);
        foreach ($exp_res_s as $exp_res) {
            if ($exp_res != '') {
                $db_func_res = DBfetch(DBselect('select itemid,name from functions where functionid=' . explode('}', $exp_res)[0]));
                $db_item_res = DBfetch(DBselect('select key_ from items where itemid=' . $db_func_res['itemid']));
                $r_expression = $r_expression . $db_func_res['name'] . '(/' . $data['export_template_data'][0]['name'] . '/' . $db_item_res['key_'] . ')' . explode('}', $exp_res)[1];
            }
        }

        if ($trigger['flags'] == 0) {
            $r_flags = 'a plain trigger';
        }
        elseif ($trigger['flags'] == 4) {
            $r_flags = 'a discovered trigger';
        }
        
        if ($trigger['type'] == 0) {
            $r_type = 'do not generate multiple events';
        }
        elseif ($trigger['type'] == 1) {
            $r_type = 'generate multiple events';
        }
        
        if ($trigger['status'] == 0) {
            $r_status = 'enabled';
        }
        elseif ($trigger['status'] == 1) {
            $r_status = 'disabled';
        }
        
        if ($trigger['state'] == 0) {
            $r_state = 'trigger state is up to date';
        }
        elseif ($trigger['state'] == 1) {
            $r_state = 'current trigger state is unknown';
        }
        
        if ($trigger['value'] == 0) {
            $r_value = 'OK';
        }
        elseif ($trigger['value'] == 1) {
            $r_value = 'problem';
        }

        if ($trigger['priority'] == 0) {
            $r_priority = 'not classified';
        }
        elseif ($trigger['priority'] == 1) {
            $r_priority = 'information';
        }
        elseif ($trigger['priority'] == 2) {
            $r_priority = 'warning';
        }
        elseif ($trigger['priority'] == 3) {
            $r_priority = 'average';
        }
        elseif ($trigger['priority'] == 4) {
            $r_priority = 'high';
        }
        elseif ($trigger['priority'] == 5) {
            $r_priority = 'disaster';
        }

        if ($trigger['recovery_mode'] == 0) {
            $r_recovery_mode = 'Expression';
        }
        elseif ($trigger['recovery_mode'] == 1) {
            $r_recovery_mode = 'Recovery expression';
        }
        elseif ($trigger['recovery_mode'] == 2) {
            $r_recovery_mode = 'None';
        }

        $r_recovery_expression = '';
        $exp_res_s = explode('{', $trigger['recovery_expression']);
        foreach ($exp_res_s as $exp_res) {
            if ($exp_res != '') {
                $db_func_res = DBfetch(DBselect('select itemid,name from functions where functionid=' . explode('}', $exp_res)[0]));
                $db_item_res = DBfetch(DBselect('select key_ from items where itemid=' . $db_func_res['itemid']));
                $r_recovery_expression = $r_recovery_expression . $db_func_res['name'] . '(/' . $data['export_template_data'][0]['name'] . '/' . $db_item_res['key_'] . ')' . explode('}', $exp_res)[1];
            }
        }

        if ($trigger['correlation_mode'] == 0) {
            $r_correlation_mode = 'All problems';
        }
        elseif ($trigger['correlation_mode'] == 1) {
            $r_correlation_mode = 'All problems if tag values match';
        }

        if ($trigger['manual_close'] == 0) {
            $r_manual_close = 'No';
        }
        elseif ($trigger['manual_close'] == 1) {
            $r_manual_close = 'Yes';
        }

        $r_opdata = $trigger['opdata'];

        $r_event_name = $trigger['event_name'];

        $r_comments = $trigger['comments'];

        $num = count($trigger['dependencies']);
        if ($num > 1) {
            $n++;
            for ($x=0; $x<$num; $x++) {
                $r_trigger_depen_descr = $trigger['dependencies'][$x]['description'];

                $d++;

                if ($x == 0) {
                    $writer->writeSheetRow($sheet6, array($n, $r_expression, $r_flags, $r_type, $r_status, $r_state, $r_value, $r_priority, $r_recovery_mode, $r_recovery_expression, $r_correlation_mode, $r_manual_close, $r_opdata, $r_trigger_depen_descr, $r_event_name, $r_comments), $row_styles);
                }
                else {
                    $writer->writeSheetRow($sheet6, array($n, '', '', '', '', '', '', '', '', '', '', '', '', $r_trigger_depen_descr, '', ''), $row_styles);
                }
            }

            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=0, $end_row=$d, $end_col=0);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=1, $end_row=$d, $end_col=1);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=2, $end_row=$d, $end_col=2);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=3, $end_row=$d, $end_col=3);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=4, $end_row=$d, $end_col=4);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=5, $end_row=$d, $end_col=5);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=6, $end_row=$d, $end_col=6);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=7, $end_row=$d, $end_col=7);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=8, $end_row=$d, $end_col=8);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=9, $end_row=$d, $end_col=9);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=10, $end_row=$d, $end_col=10);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=11, $end_row=$d, $end_col=11);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=12, $end_row=$d, $end_col=12);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=14, $end_row=$d, $end_col=14);
            $writer->markMergedCell($sheet6, $start_row=$d-($num-1), $start_col=15, $end_row=$d, $end_col=15);
        }
        elseif ($num == 1) {
            $n++;
            $d++;
            $r_trigger_depen_descr = $trigger['dependencies'][0]['description'];
            $writer->writeSheetRow($sheet6, array($n, $r_expression, $r_flags, $r_type, $r_status, $r_state, $r_value, $r_priority, $r_recovery_mode, $r_recovery_expression, $r_correlation_mode, $r_manual_close, $r_opdata, $r_trigger_depen_descr, $r_event_name, $r_comments), $row_styles);
        }
        else {
            $n++;
            $d++;
            $writer->writeSheetRow($sheet6, array($n, $r_expression, $r_flags, $r_type, $r_status, $r_state, $r_value, $r_priority, $r_recovery_mode, $r_recovery_expression, $r_correlation_mode, $r_manual_close, $r_opdata, '', $r_event_name, $r_comments), $row_styles);
        }
    }
}

$host_proto_header = array(
  'no.'=>'@',
  'name'=>'@',
  'status'=>'@',
  'discover'=>'@',
  'custom_interfaces'=>'@',
  'group'=>'@'
);

$sheet7 = 'host proto';

$writer->writeSheetHeader($sheet7, $host_proto_header, array_merge($header_styles, ['widths'=>[5,30,20,30,30,20]]));

$p = 0;
foreach ($data['export_host_proto_data'] as $hosts) {
    foreach ($hosts as $host) {
        $r_name = $host['name'];

        if ($host['status'] == 0) {
            $r_status = 'monitored host';
        }
        elseif ($host['status'] == 1) {
            $r_status = 'unmonitored host';
        }
        
        if ($host['discover'] == 0) {
            $r_discover = 'new hosts will be discovered';
        }
        elseif ($host['discover'] == 1) {
            $r_discover = 'new hosts will not be discovered and existing hosts will be marked as lost';
        }
        
        if ($host['custom_interfaces'] == 0) {
            $r_custom_interfaces = 'inherit interfaces from parent host';
        }
        elseif ($host['custom_interfaces'] == 1) {
            $r_custom_interfaces = 'use host prototypes custom interfaces';
        }
        
        $db_group_res = DBfetch(DBselect('select name from hstgrp where groupid=' . $host['groupLinks'][0]['groupid']));
        $r_group = $db_group_res['name'];

        $p++;

        $writer->writeSheetRow($sheet7, array($p, $r_name, $r_status, $r_discover, $r_custom_interfaces, $r_group), $row_styles);
    }
}

$writer->writeToStdOut();