<?php

//Таблица с данными
$xlsx = 'СИоЗП_2021-03-15.xlsx';
//Рег.номер в ПФР
$pfrnumber = '111-111-111111';
//ИНН
$inn = '0123456789';
//КПП
$kpp='123456789';
//ОКФС
$okfs='13';
//КТО
$kto='4.1.01';
//ОКОГУ
$okogu='1320700';

function toMoney(string $value)
{
	$num = preg_split("/\./", $value);
	return $num[0].'.'.str_pad($num[1],2,'0',STR_PAD_RIGHT);
}

function toStagWithMonth(string $value)
{
	$stages = preg_split("/\./", $value);
	return str_pad($stages[0],2,'0',STR_PAD_LEFT).'.'.str_pad($stages[1],2,'0',STR_PAD_LEFT);
}

function toStagRound(string $value)
{
	$stages = preg_split("/\./", $value);
	return $stages[0];
}

function GUID()
{
	return sprintf('%04X%04X-%04X-%04X-%04X-%04X%04X%04X', mt_rand(0, 65535), mt_rand(0, 65535), mt_rand(0, 65535), mt_rand(16384, 20479), mt_rand(32768, 49151), mt_rand(0, 65535), mt_rand(0, 65535), mt_rand(0, 65535));
}

$xml_out = simplexml_load_file('template_document.xml');

$zip = new ZipArchive;
$res = $zip->open($xlsx);

if ($res === TRUE) {
	$xml_shared_strings = simplexml_load_string($zip->getFromName('xl/sharedStrings.xml'));

	$sheet1 = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
	foreach ($sheet1->sheetData->children() as $row)
	{
		if ($row->attributes()->r>=6)
		{
			$xml_out->СИоЗП->СЗП->addChild('Период');
			$num_period = count($xml_out->СИоЗП->СЗП->Период)-1;
			foreach ($row->children() as $c)
			{
				$column = preg_split("/\d+/", $c->attributes()->r)[0];
				$num_si=(int)$c->v->__toString();
				
				if (isset($c->attributes()->t[0]))
				{
					$val = 'ОШИБКА';
					if ($c->attributes()->t[0] =='s')
					{
						$val = str_replace(',','.',$xml_shared_strings->children()->si[$num_si]->t->__toString());
					}
					
				}
				else
				{
					$val = $num_si;
				}
				
				if ($val!='0.00')
				{
				switch($column)
				{
					case 'A':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->addChild('ОтчетныйПериод');
						$xml_out->СИоЗП->СЗП->Период[$num_period]->ОтчетныйПериод->addChild('Год',$val);
						break;
					case 'B':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->ОтчетныйПериод->addChild('Месяц',$val);
						break;
					case 'G':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->addChild('Работник');
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->addChild('ФИО');
						$fio = preg_split("/\s+/", $val,3);
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->ФИО->addChild('Фамилия',$fio[0]);
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->ФИО->addChild('Имя',$fio[1]);
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->ФИО->addChild('Отчество',$fio[2]);
						break;
					case 'H':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->addChild('СНИЛС',$val);
						break;
					case 'I':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->addChild('ОбщийСтаж',toStagRound($val));
						break;
					case 'J':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->addChild('СЗПД');
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('Должность',$val);
						break;
					case 'K':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ККП',$val);
						break;
					case 'L':
						switch(strtolower($val))
						{
							case 'основное':
								$val = '1';
								break;
							case 'внешнее совместительство':
								$val = '2';
								break;
							case 'внутреннее совместительство':
								$val = '3';
								break;
							default:
								$val = '';
						}
						if ($val!='')
						{
							$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('УсловиеЗанятости',$val);
						}
						break;
					case 'M':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('Ставка',$val);
						break;
					case 'N':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('РабВремяНорма',$val);
						break;
					case 'O':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('РабВремяФакт',$val);
						break;
					case 'P':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('НачисленияТариф',toMoney($val));
						break;
					case 'Q':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ОУТ',$val);
						break;
					case 'R':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('НачисленияОУТ',toMoney($val));
						break;
					case 'S':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ДоплатаСовмещение',toMoney($val));
						break;
					case 'T':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('НачисленияИныеФед',toMoney($val));
						break;
					case 'U':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('НачисленияИныеРег',toMoney($val));
						break;
					case 'V':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('НачисленияПремии',toMoney($val));
						break;
					case 'W':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('НепрерывныйСтаж',toStagWithMonth($val));
						break;
					case 'X':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ДоплатаСтаж',toMoney($val));
						break;
					case 'Y':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ДоплатаСМ',toMoney($val));
						break;
					case 'Z':
						switch(strtolower($val))
						{
							case 'первая':
								$val = '1';
								break;
							case 'вторая':
								$val = '2';
								break;
							case 'высшая':
								$val = '3';
								break;
							default:
								$val = '';
						}
						if ($val!='')
						{
							$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('КвалКатегория',$val);
						}
						break;
					case 'AA':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ДоплатаКвалКат',toMoney($val));
						break;
					case 'AB':
						switch(strtolower($val))
						{
							case 'кандидат наук':
								$val = '1';
								break;
							case 'доктор наук':
								$val = '2';
								break;
							default:
								$val = '';
						}
						if ($val!='')
						{
							$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('КвалКатегория',$val);
						}
						
						break;
					case 'AC':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ДоплатаУС',toMoney($val));
						break;
					case 'AD':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ДоплатаНаставничество',toMoney($val));
						break;
					case 'AE':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ДоплатаМолодСпец',toMoney($val));
						break;
					case 'AF':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ВыплатыИныеСтимул',toMoney($val));
						break;
					case 'AG':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ВыплатыПрочие',toMoney($val));
						break;
					case 'AH':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('ВыплатыКомпенс',toMoney($val));
						break;
					case 'AI':
						$xml_out->СИоЗП->СЗП->Период[$num_period]->Работник->СЗПД->addChild('НачисленияИтого',toMoney($val));
						break;
				} //end_switch
			} //endif
				
			}//column
		}//endif
	}//row

	$sheet2 = simplexml_load_string($zip->getFromName('xl/worksheets/sheet2.xml'));
	foreach($sheet2->sheetData->children()[4] as $c)
	{
		$column = preg_split("/\d+/", $c->attributes()->r)[0];
		$num_si=(int)$c->v->__toString();
		$val = str_replace(',','.',$xml_shared_strings->children()->si[$num_si]->t->__toString());
		switch($column)
		{
			case 'A':
				$xml_out->СИоЗП->ФондЗП->Период->Год[0]=$val;
				break;
			case 'E':
				$xml_out->СИоЗП->ФондЗП->Период->РасхОбщФед[0]=toMoney($val);
				break;
			case 'F':
				$xml_out->СИоЗП->ФондЗП->Период->РасхКатФед[0]=toMoney($val);
				break;
			case 'G':
				$xml_out->СИоЗП->ФондЗП->Период->РасхОбщСуб[0]=toMoney($val);
				break;
			case 'H':
				$xml_out->СИоЗП->ФондЗП->Период->РасхКатСуб[0]=toMoney($val);
				break;
			case 'I':
				$xml_out->СИоЗП->ФондЗП->Период->РасхОбщМун[0]=toMoney($val);
				break;
			case 'J':
				$xml_out->СИоЗП->ФондЗП->Период->РасхКатМун[0]=toMoney($val);
				break;
			case 'K':
				$xml_out->СИоЗП->ФондЗП->Период->РасхОбщОМС[0]=toMoney($val);
				break;
			case 'L':
				$xml_out->СИоЗП->ФондЗП->Период->РасхКатОМС[0]=toMoney($val);
				break;
			case 'M':
				$xml_out->СИоЗП->ФондЗП->Период->РасхОбщИные[0]=toMoney($val);
				break;
			case 'N':
				$xml_out->СИоЗП->ФондЗП->Период->РасхКатИные[0]=toMoney($val);
				break;
		} //endswitch
		
	} //endforeach
	
	$sheet3 = simplexml_load_string($zip->getFromName('xl/worksheets/sheet3.xml'));
	foreach($sheet3->sheetData->children()[3] as $c)
	{
		$column = preg_split("/\d+/", $c->attributes()->r)[0];
		$num_si=(int)$c->v->__toString();
		$val = str_replace(',','.',$xml_shared_strings->children()->si[$num_si]->t->__toString());
		switch($column)
		{
			case 'A':
				$xml_out->СИоЗП->СЗПРук->Период->Год[0]=$val;
				break;
			case 'D':
				$xml_out->СИоЗП->СЗПРук->Период->СредЗПРук[0]=toMoney($val);
				break;
			case 'E':
				$xml_out->СИоЗП->СЗПРук->Период->СредЗПЗам[0]=toMoney($val);
				break;
			case 'F':
				$xml_out->СИоЗП->СЗПРук->Период->СредЗПГлБух[0]=toMoney($val);
				break;
			case 'G':
				$xml_out->СИоЗП->СЗПРук->Период->СредЗПРаб[0]=toMoney($val);
				break;
		} //endswitch
		
	} //endforeach
	
	$guid = GUID();
	$xml_out->СлужебнаяИнформация->GUID[0] = $guid;
	$xml_out->СлужебнаяИнформация->ДатаВремя[0] = (string)date(DATE_ATOM);

	$xml_out->СИоЗП->Организация->ИНН[0] = $inn;
	$xml_out->СИоЗП->Организация->КПП[0] = $kpp;
	$xml_out->СИоЗП->Организация->ОКФС[0] = $okfs;
	$xml_out->СИоЗП->Организация->КТО[0] = $kto;
	$xml_out->СИоЗП->Организация->ОКОГУ[0] = $okogu;
	
	$dom = new DOMDocument("1.0");
	$dom->preserveWhiteSpace = false;
	$dom->formatOutput = true;
	$dom->loadXML($xml_out->asXML());
	$dom->save('xml_out_report.xml');
	
	$content = file_get_contents('xml_out_report.xml');
	$word_maps = [
    'pf.rf/siozp' => 'пф.рф/СИоЗП',
    'pf.rf/ut' => 'пф.рф/УТ',
	'pf.rf/af' => 'пф.рф/АФ',
	'ФИО' => 'УТ2:ФИО',
	'СНИЛС' => 'УТ2:СНИЛС',
	'Фамилия' => 'УТ2:Фамилия',
	'Имя' => 'УТ2:Имя',
	'Отчество' => 'УТ2:Отчество',
	'GUID' => 'АФ5:GUID',
	'ДатаВремя' => 'АФ5:ДатаВремя',
	'ИНН' => 'УТ2:ИНН',
	'КПП' => 'УТ2:КПП'
	];
	
	file_put_contents('ПФР_'.preg_split("/-/", $pfrnumber)[0].preg_split("/-/", $pfrnumber)[1].'_СИоЗП_'.preg_split("/-/", $pfrnumber)[0].preg_split("/-/", $pfrnumber)[1].preg_split("/-/", $pfrnumber)[2].'_'.(string)date('Ymd').'_'.$guid.'.xml', strtr($content, $word_maps));

/* 	$dom->load('template_period.xml');
	$tst = new DOMDocument;
	$tst->loadXML('<root><element><child>1</child><child>2</child></element></root>');
	$node = $tst->getElementsByTagName('element')->item(0);
	$node = $dom->importNode($node, true);
	$dom->getElementsByTagName('ФИО')->item(0)->appendChild($node);
	//$dom->getElementsByTagName('ФИО')[0]->appendChild($node);
	$dom->save('test.xml'); */
	$zip->close();
}

