<?php

date_default_timezone_set("America/New_York");

define('EOL', (PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once(__DIR__ . '/../../../global_includes/classes/PHPExcel.php');
include_once(__DIR__ . '/../../classes/quote.php');
include_once(__DIR__ . '/../../classes/item.php');
include_once(__DIR__ . '/../../classes/agency.php');
include_once(__DIR__ . '/../../classes/invoice.php');
include_once(__DIR__ . '/../../includes/date_functions.php');
include_once(__DIR__ . '/../../includes/sorting_functions.php');
include_once(__DIR__ . '/../../classes/currency.php');
include_once(__DIR__ . '/../../classes/inv_item_cx_policy.php');


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$data = array();

// Set value binder
PHPExcel_Cell::setValueBinder(new PHPExcel_Cell_AdvancedValueBinder ());

$invoice_id = (isset ($_POST['invoice_id'])) ? $_POST['invoice_id'] : $_GET['invoice_id'];
$withAuthorizationForm = (isset ($_GET['withAuthorizationForm'])) ? $_GET['withAuthorizationForm'] : false;

$cancellation_policies = array();

$currency = new Currency();

$total_agency_commission = 0;

$secure_message = "All travel arrangements outlined on this invoice have been secured";
// This array contains a key with the currency code selected and an array of totals inside it.
// The USD array contains the sum of all totals as it comes from the database for display at the bottom while the other currencies each have their own totals
// i.e:
// [CAD] {total_retail:450.43}
// [EUR] {total_retail:422.00}
// [USD] {total_retail:790.33} This is the sum of all values NOT JUST USD
$currencies_array = array();
$currencies_array['USD']['total_retail'] = 0;


$invoice = new Invoice ($invoice_id);
$invoice_number = $invoice->GetData()->invoice_number;   // This needs to be modified

$agency_info = $invoice->GetAgencyInfo();

// INSTANTIATE QUOTE ITEM
$quote = new Quote ($invoice->data->quote_id);
$quote_data = $quote->GetQuote();

// GET ALL RELATED ITEMS
$item_ids = $invoice->GetModifiedItemIds();

// Determine if this is a credit note or an invoice
if ($invoice->data->amount_invoiced > 0) {
	$invoice_type = 'INVOICE';
} elseif ($invoice->data->amount_invoiced < 0) {
	$invoice_type = 'CREDIT NOTE';
} else {
	$invoice_type = 'INTERNAL MODIFICATION';
}


// STRUCTURE ITEMS AND ADD THEM TO OUR INPUT
foreach ($item_ids as $key => $item_id) {

	$item = new InvoiceItem ($item_id);
	$item->SetSupplierInfo();

	$supplier_currency = $item->reservation['rates_data']['supplier_currency'];

	$cur_mod_type = $item->data->modification_type;
	$prv_mod_type = $item->data->modification_type_previous;

	$include_item = true;

	// Parse all the scenarios where the internal modifications should not be included in the invoice
	switch ($cur_mod_type) {
		case 'INT-MOD':
		case 'INT-MOD-CR':
			$include_item = false;
			break;
		case 'MODIFIED-INT':
		case 'MODIFIED':
		case 'CREDITED':
			if (($prv_mod_type == 'INT-MOD') || ($prv_mod_type == 'MODIFIED-INT')) {
				$include_item = false;
			}
			break;
	}

	if ($include_item) {

		// Remove items not marked as visible in quote / ib
		$item_data = $item->reservation;
		$item_data['modification_type'] = $item->data->modification_type;
		$item_data['reservation_type'] = $item->data->reservation_type;
		$item_data['check_in_date'] = $item->data->check_in_date;
		$item_data['check_out_date'] = $item->data->check_out_date;

		$policy = new InvItemCxPolicy ($item_id, $item->data->check_in_date, $supplier_currency);

		$item_data['rates_data']['policy'] = $policy->get_policy();
		$data['items'][] = $item_data;

		$supplier_currency = $item_data['rates_data']['supplier_currency'];
		$item_exchange_rate = $item_data['rates_data']['exchange_rate'];
		$units = $item_data['request']['units'];

		$total_retail_usd = $item_data['rates_data']['total']['rate_retail_after_tax'] * $units;
		$total_retail = $currency->USDToCurrency($item_data['rates_data']['total']['rate_retail_after_tax'] * $units, $supplier_currency, $item_exchange_rate);


		// Get the totals to display at the bottom
		// Check if the current currency already exists. If it doesn't, initialize it.
		if (!isset ($currencies_array[$supplier_currency])) {
			$currencies_array[$supplier_currency]['total_retail'] = 0;
		}
		// Update this currency's total, if different that USD which is updated by default
		if ($supplier_currency != 'USD') {
			$currencies_array[$supplier_currency]['total_retail'] += $total_retail;
		}

		// Update USD totals by default
		$currencies_array['USD']['total_retail'] += $total_retail_usd;
	} else {
		// This is an internal modification item, remove from the invoice
		unset ($item_ids[$key]);
	}
}


// Sort the items in this quote by their position
if (isset($data['items']) && count($data['items']) > 0) {
	uasort($data['items'], 'sort_items_by_position');
} else {
	$data['items'] = array();
}


// Create the XML document and put information on it
// Set document properties
$objPHPExcel->getProperties()->setCreator("Agent Name")
	->setLastModifiedBy("Agent Name")
	->setTitle("OTI Invoice")
	->setSubject("OTI Invoice")
	->setDescription("OTI Invoice")
	->setKeywords("")
	->setCategory("");

// Display the logo
$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setName('Logo');
$objDrawing->setDescription('Logo');
//$objDrawing -> setPath ( __DIR__ . '/../../images/logo_overseasinternational_blk.png' );
$objDrawing->setPath(__DIR__ . '/../../images/olg_logo_new.png');
$objDrawing->setHeight(82);
$objDrawing->setWidth(82);
$objDrawing->setOffsetY(5);
$objDrawing->setCoordinates('E1');

$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());


// Set column widths
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(3);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(10);

// Invoice box
SetBorder($objPHPExcel, 'A4:C7', 'MEDIUM');
SetBold($objPHPExcel, 'A5');
SetBold($objPHPExcel, 'A6');

// Logo row
$objPHPExcel->getActiveSheet()->mergeCells('A1:I1');
SetHorizontalAlignment($objPHPExcel, 'A1', 'CENTER');
$objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(90);


// Invoice headline
$objPHPExcel->getActiveSheet()->mergeCells('A2:I2');
SetBorder($objPHPExcel, 'A2:I2', 'MEDIUM');
SetBold($objPHPExcel, 'A2');
SetHorizontalAlignment($objPHPExcel, 'A2', 'CENTER');
$objPHPExcel->getActiveSheet()->getStyle('A2')->getFont()->setSize(14);

//SECURITY MESSAGE
$objPHPExcel->getActiveSheet()->mergeCells('A3:I3');
SetBorder($objPHPExcel, 'A3:I3', 'MEDIUM');
SetHorizontalAlignment($objPHPExcel, 'A3', 'CENTER');
$objPHPExcel->getActiveSheet()->setCellValue('A3', $secure_message)->getRowDimension(3)->setRowHeight(20);

// Agency info panel
SetHorizontalAlignment($objPHPExcel, 'C6', 'RIGHT');
SetBorder($objPHPExcel, 'E4:I7', 'MEDIUM');
$objPHPExcel->getActiveSheet()->mergeCells('F4:I4');
$objPHPExcel->getActiveSheet()->mergeCells('F5:I5');
$objPHPExcel->getActiveSheet()->mergeCells('F6:I6');
$objPHPExcel->getActiveSheet()->mergeCells('F7:I7');
SetBold($objPHPExcel, 'E4');
SetBold($objPHPExcel, 'E5');
SetBold($objPHPExcel, 'E6');
SetBold($objPHPExcel, 'E7');

$objPHPExcel->getActiveSheet()
	->getStyle('C5')
	->getNumberFormat()
	->setFormatCode(
		PHPExcel_Style_NumberFormat::FORMAT_TEXT
	);

$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue('E4', 'Client:')
	->setCellValue('E5', 'Agency:')
	->setCellValue('E6', 'Agent:')
	->setCellValue('E7', 'Location:')
	->setCellValue('C5', "$invoice_number ")
	->setCellValue('C6', date('M d, Y', strtotime($item->data->creation_date)))
	->setCellValue('F4', $quote_data['client_name'])
	->setCellValue('F5', $agency_info['agency_name'])
	->setCellValue('F6', $quote_data['agent_name'])
	->setCellValue('F7', $agency_info['agency_country']);


// Add some data
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue('A2', $invoice_type)
//    ->setCellValue('A3', $secure_message)
	->setCellValue('A4', $invoice_type)
	->setCellValue('A5', 'Number:')
	->setCellValue('A6', 'Date:');



// Row 11
SetHorizontalAlignment($objPHPExcel, 'F11', 'RIGHT');
SetHorizontalAlignment($objPHPExcel, 'G11', 'RIGHT');
SetHorizontalAlignment($objPHPExcel, 'H11', 'RIGHT');
SetHorizontalAlignment($objPHPExcel, 'I11', 'RIGHT');
$objPHPExcel->getActiveSheet()->mergeCells('A11:E11');
SetBold($objPHPExcel, 'A8');
SetBold($objPHPExcel, 'A9');
SetBold($objPHPExcel, 'A10');
SetBold($objPHPExcel, 'A11');
SetBold($objPHPExcel, 'F11');
SetBold($objPHPExcel, 'G11');
SetBold($objPHPExcel, 'H11');
SetBold($objPHPExcel, 'I11');
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue('A11', 'Description of Services')
	->setCellValue('F11', 'Unit Price')
	->setCellValue('G11', 'Nights')
	->setCellValue('H11', 'Qty.')
	->setCellValue('I11', 'Subtotal');


$objPHPExcel->getActiveSheet()->setCellValue('A8', "On your behalf, we are donating 10 USD to Sandy Hook Promise.")->getRowDimension(8)->setRowHeight(20);
$objPHPExcel->getActiveSheet()->setCellValue('A9', "Read more here: https://www.sandyhookpromise.org/our_impact")->getRowDimension(8)->setRowHeight(20);

// Spacer rows
//$objPHPExcel->getActiveSheet()->getRowDimension(3)->setRowHeight(5);
$objPHPExcel->getActiveSheet()->setCellValue('A10', 'Please note our bank information has changed.  Details noted at bottom of invoice.')->getRowDimension(8)->setRowHeight(20);
$objPHPExcel->getActiveSheet()->getStyle('A10')->applyFromArray(
	array(
		'fill' => array(
			'type' => PHPExcel_Style_Fill::FILL_SOLID,
			'color' => array('rgb' => 'FFFF00')
		)
	)
);
//$objPHPExcel -> getActiveSheet () ->setCellValue('A8','')-> getRowDimension ( 8 ) -> setRowHeight ( 20 );
$objPHPExcel->getActiveSheet()->getRowDimension(12)->setRowHeight(5);


//$objPHPExcel->getActiveSheet()->getStyle('A8:J8')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->mergeCells('A8:I8');
$objPHPExcel->getActiveSheet()->mergeCells('A9:I9');
$objPHPExcel->getActiveSheet()->mergeCells('A10:I10');

// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle('Invoice #');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);


$row = 14; // row counter start

foreach ($data['items'] as $item) {

	$request = $item['request'];
	$rates_data = $item['rates_data'];
	$daily_rates = $rates_data['daily_rates'];

	$units = $request['units'];

	// Get the suppliers currency
	$supplier_currency = $item['rates_data']['supplier_currency'];
	$item_exchange_rate = $item['rates_data']['exchange_rate'];

	// Build the item name depending on the source and Store the cancellation policy
	// for each item for later use at the bottom of the invoice.
	switch ($item['reservation_type']) {
		case 'hotel':
			$name = 'Hotel: ' . $item['hotel_info']['hotel_name'];
			break;
		case 'activity':
			$name = 'Activity: ' . $item['activity_info']['activity_name'];
			break;
		case 'villa':
			$name = 'Villa: ' . $item['activity_info']['activity_name'];
			break;
		case 'transfer':
			$name = 'Transfer - ' . $item['pickup_info']['pickup_name'] . " to " . $item['dropoff_info']['dropoff_name'];
			break;
		case 'resort_fee':
			$name = 'Resort Fee: ' . $item['hotel_info']['hotel_name'] . ' - RF';
			break;
		case 'hotel_service':
			$name = 'Hotel Service: ' . $item['hotel_info']['hotel_name'] . ' - ' . $item['service_info']['service_name'];
			break;
		case 'car_rental':
			$name = 'Car Rental - ' . $item['supplier_info']['supplier_name'] . ' - ' . $item['request']['vehicle_group'];
			break;
		case 'taxable_fee':
			if (!isset ($item['taxable_fee_info']['taxable_fee_name'])) {
				$item['taxable_fee_info']['taxable_fee_name'] = 'Fee';
			}

			$name = 'Activity Fee: ' . $item['activity_info']['activity_name'] . ' - ' . $item['taxable_fee_info']['taxable_fee_name'];
			break;
	}

	// In the case of a modification, the item is credited at the original rate and added back at thenew rate. Show it here in the items name
	if ($item['modification_type'] == 'MOD-CR') {
		$name .= " (Credited for Modification)";
	}


	// Build the cancellation policies strings for each item. (each item might have more than one cancellation policy)
	$cx_policies = $item['rates_data']['policy'];

	if (($item['reservation_type'] == 'hotel') || ($item['reservation_type'] == 'activity') || ($item['reservation_type'] == 'transfer') || ($item['reservation_type'] == 'car_rental')) {
		// Do not show cancellation policies for credited items or mod-cr
		if (($item['modification_type'] != 'MOD-CR') && ($item['modification_type'] != 'CREDIT')) {

			foreach ($cx_policies as $cx_policy_key => $cx_policy) {
				$cancellation_policies[] = $name . ': ' . $cx_policy['policy_description'];
			}
		}
	}


	$active_cell = 'A' . $row;

	// Write Item name
	$objPHPExcel->getActiveSheet()->mergeCells($active_cell . ':I' . $row);
	SetBold($objPHPExcel, $active_cell);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, $name);

	// Write 'dates' once
	$active_cell = 'A' . ($row + 1);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'Dates:');

	// Getting intervals
	for ($i = 0; $i < count($daily_rates); $i++) {
		$wrap_it_up = false; // interval end flag
		$same_day_booking = FALSE;

		if ($item['check_in_date'] == $item['check_out_date']) {
			$wrap_it_up = TRUE;
			$same_day_booking = TRUE;
		}

		$rate_retail_before_tax_usd = $daily_rates[$i]['rate_retail_before_tax'];

		$rate_retail_before_tax = $currency->USDToCurrency($rate_retail_before_tax_usd, $supplier_currency, $item_exchange_rate);

		$current_date = $daily_rates[$i]['rate_date'];
		if (!isset ($interval_start)) {
			$interval_start = $current_date;
		}

		// Check if there is one more day of rates
		if (isset ($daily_rates[$i + 1])) {
			// IF TODAY'S RATE_COMMISSION_NET DOESNT EQUAL TOMORROWS
			if ($daily_rates[$i]['rate_commission_percent_net'] != $daily_rates[$i + 1]['rate_commission_percent_net']) {
				$wrap_it_up = true;
			} // IF TODAY'S RATE_RETAIL DOESNT EQUAL TOMORROWS
			else if ($daily_rates[$i]['rate_retail_before_tax'] != $daily_rates[$i + 1]['rate_retail_before_tax']) {
				$wrap_it_up = true;
			} // IF TODAY'S RATE_FEE DOESNT EQUAL TOMORROWS
			else if ($daily_rates[$i]['rate_fee'] != $daily_rates[$i + 1]['rate_fee']) {
				$wrap_it_up = true;
			}
		} else {
			$wrap_it_up = true;
		}

		$interval_end = $current_date;

		if ($wrap_it_up == true) {

			// get the number of days / nights
			$end_stamp = new DateTime ($interval_end);
			$start_stamp = new DateTime ($interval_start);
			$interval = $end_stamp->diff($start_stamp);

			$days = ($interval->days) + 1;

			// Output
			$row++; // row counter
			// Write the date interval
			$active_cell = 'B' . $row;

			if ($same_day_booking) {
				$cell_value = date("M d, Y", strtotime($interval_start));
				$objPHPExcel->setActiveSheetIndex(0)
					->setCellValue($active_cell, $cell_value);
			} else {
				if (($item['reservation_type'] == 'hotel') || ($item['reservation_type'] == 'resort_fee')) {
					$cell_value = date("M d, Y", strtotime($interval_start)) . ' to ' . date("M d, Y", strtotime($interval_end) + 86400);
					$objPHPExcel->setActiveSheetIndex(0)
						->setCellValue($active_cell, $cell_value);
				} else {
					$interval_end = date("M d, Y", strtotime($interval_end . ' - 1 day'));
					$cell_value = date("M d, Y", strtotime($interval_start)) . ' to ' . date("M d, Y", strtotime($interval_end) + 86400);
					$objPHPExcel->setActiveSheetIndex(0)
						->setCellValue($active_cell, $cell_value);
				}
			}

			// Write the supplier's currency
			$active_cell = 'E' . $row;
			SetHorizontalAlignment($objPHPExcel, $active_cell, 'RIGHT');
			$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue($active_cell, "($supplier_currency)");

			// Write the Unit Price
			$active_cell = 'F' . $row;
			$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getNumberFormat()
				->setFormatCode('#,##0.00');
			$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue($active_cell, $rate_retail_before_tax);

			// Write the number of nights
			$active_cell = 'G' . $row;
			$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue($active_cell, $days);

			// Write the number of units
			$active_cell = 'H' . $row;
			$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue($active_cell, $units);

			// Write the subtotal
			$active_cell = 'I' . $row;
			$cell_value = '=F' . $row . '*G' . $row . '*H' . $row;
			$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getNumberFormat()
				->setFormatCode('#,##0.00');
			$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue($active_cell, $cell_value);


			unset ($interval_start);
		}
	}


	if ($item['reservation_type'] == 'hotel') {

		$row++;    // go to the next row
		// Write the room type
		$active_cell = 'A' . $row;
		$objPHPExcel->getActiveSheet()->mergeCells($active_cell . ':E' . $row);
		$cell_value = 'Room Category: ' . $item['room_info']['room_name'];
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $cell_value);

		$row++;    // go to the next row
		// Write the bedding type
		$active_cell = 'A' . $row;
		$objPHPExcel->getActiveSheet()->mergeCells($active_cell . ':E' . $row);
		$cell_value = 'Bedding: ' . $item['room_info']['room_beds'];
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $cell_value);
	}

	if ($item['reservation_type'] == 'car_rental') {

		$row++;    // go to the next row
		// Ovewrite the date above with real pickup and dropoff dates
		$active_cell = 'B' . ($row - 1);
		$cell_value = date('M d, Y', strtotime($item['check_in_date'])) . ' to ' . date('M d, Y', strtotime($item['check_out_date']));
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $cell_value);

		// Write the notes
		$active_cell = 'A' . $row;
		$objPHPExcel->getActiveSheet()->mergeCells($active_cell . ':E' . $row);
		$cell_value = $item['request']['notes'];
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $cell_value);
	}


	if ($item['reservation_type'] == 'transfer') {

		$row++;  // go to the next row
		// Write the vehicle type
		$active_cell = 'A' . $row;
		$objPHPExcel->getActiveSheet()->mergeCells($active_cell . ':E' . $row);
		$cell_value = 'Car Category: ' . $item['vehicle_info']['vehicle_name'];
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $cell_value);
	}


	if ($item['reservation_type'] == 'activity') {

		if ($item['activity_info']['notes1'] != '') {
			$row++;  // go to the next row
			//
			$active_cell = 'A' . $row;
			$objPHPExcel->getActiveSheet()->mergeCells($active_cell . ':E' . $row);
			$cell_value = $item['activity_info']['notes1'];
			$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue($active_cell, $cell_value);
		}
		if ($item['activity_info']['notes2'] != '') {
			$row++;  // go to the next row
			//
			$active_cell = 'A' . $row;
			$objPHPExcel->getActiveSheet()->mergeCells($active_cell . ':E' . $row);
			$cell_value = $item['activity_info']['notes2'];
			$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue($active_cell, $cell_value);
		}
		if ($item['activity_info']['notes3'] != '') {
			$row++;  // go to the next row
			//
			$active_cell = 'A' . $row;
			$objPHPExcel->getActiveSheet()->mergeCells($active_cell . ':E' . $row);
			$cell_value = $item['activity_info']['notes3'];
			$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue($active_cell, $cell_value);
		}
	}


	// Check if there are any taxes and display accordingly
	if ($item['rates_data']['total']['rate_tax'] > 0) {

		$row++;    // go to the next row
		// Check if there are taxes and display them
		$active_cell = 'D' . $row;
		$cell_value = 'Taxes: ' . number_format($item['rates_data']['total']['rate_tax_percent'], 2) . '%';
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $cell_value);

		// Show the taxes per night
		$active_cell = 'F' . $row;
		$rate_tax_usd = $item['rates_data']['total']['rate_tax'];

		$cell_value = $currency->USDToCurrency($rate_tax_usd, $supplier_currency, $item_exchange_rate);

		$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getNumberFormat()
			->setFormatCode('#,##0.00');
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $cell_value);

		// Write the number of units
		$active_cell = 'H' . $row;
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $units);

		// Show the tax subtotal
		$active_cell = 'I' . $row;
		$active_cell = 'I' . $row;
		$cell_value = '=F' . $row . '*H' . $row;
		$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getNumberFormat()
			->setFormatCode('#,##0.00');
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $cell_value);
	}

	// Check if there are any fees and display accordingly
	if ($item['rates_data']['total']['rate_fee'] > 0) {
		$row++; // Advance to the next row
		// Show the fees
		$active_cell = 'D' . $row;
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, 'Fees per night:');

		// Display the fees per night
		$active_cell = 'F' . $row;
		$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getNumberFormat()
			->setFormatCode('#,##0.00');
		// Prevent divide by zero
		$item['request']['number_of_nights'] = ($item['request']['number_of_nights'] == 0) ? 1 : $item['request']['number_of_nights'];
		$fees_usd = $item['rates_data']['total']['rate_fee'] / $item['request']['number_of_nights'];
		$cell_value = $currency->USDToCurrency($fees_usd, $supplier_currency, $item_exchange_rate);

		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $cell_value);

		// Write the number of nights
		$active_cell = 'G' . $row;
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $item['request']['number_of_nights']);

		// Write the number of units
		$active_cell = 'H' . $row;
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $units);

		// Write the fees subtotal
		$active_cell = 'I' . $row;
		$cell_value = '=F' . $row . '*G' . $row . '*H' . $row;
		$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getNumberFormat()
			->setFormatCode('#,##0.00');
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $cell_value);
	}


	$total_agency_commission += ($item['rates_data']['total']['rate_commission'] * $item['request']['units']);

	// Leave a blank row in between items;
	$row++;
	$row++;    // go to the next row
}

// Put the USD section of the totals at the bottom of the array
$buffer = $currencies_array['USD'];
unset ($currencies_array['USD']);
$currencies_array['USD'] = $buffer;

// Start the loop for all the currency totals
foreach ($currencies_array as $currency_code => $totals) {

	// Display totals
	$active_cell = 'F' . $row;
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'Subtotal in:');

	// Show the currency of the invoice
	$active_cell = 'G' . $row;
	$cell_value = $currency_code;
	SetHorizontalAlignment($objPHPExcel, $active_cell, 'CENTER');
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, $cell_value);

	// Show the subtotal
	$active_cell = 'I' . $row;
	$cell_value = $totals['total_retail'];
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getNumberFormat()
		->setFormatCode('#,##0.00');
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, $cell_value);

	// Since the USD subtotal will be the last row always, keep this count so we can substract the commissions etc later on
	$usd_subtotal_row = $row;

	$row++;
}

$row++; // Advance to next row
// Display total agency commisison
$active_cell = 'F' . $row;
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, 'Commission:');

$active_cell = 'I' . $row;
$cell_value = $total_agency_commission;
$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getNumberFormat()
	->setFormatCode('#,##0.00');
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, $cell_value);


$row++; // Advance to the next row
// Calculate wire fees

switch ($invoice->data->payment_type) {
	case 'wire_transfer':
		$fee_title = 'Wire Fees ';
		break;
	case 'credit_card':
		$fee_title = "CC Fees ";
		break;
	case 'paypal':
		$fee_title = "Paypal Fees ";
		break;
	case 'cash':
		$fee_title = 'Fees ';
		break;
	case 'check':
		$fee_title = 'Check Fees ';
		break;
	case 'western_union':
		$fee_title = 'Western Union Fees ';
		break;
	case 'money_order':
		$fee_title = 'Money Order Fees ';
		break;
	case 'bank_deposit':
		$fee_title = 'Bank Deposit Fees ';
		break;
}

if ($invoice->data->payment_fee_type == 'amount') {
	$wire_fee = $invoice->data->payment_fee;
	$fee_title .= "(USD)";
} else {
	$wire_fee = '=(I' . ($usd_subtotal_row) . '-I' . ($usd_subtotal_row + 2) . ') * ' . $invoice->data->payment_fee . '/100';
	$fee_title .= "({$invoice -> data -> payment_fee}%)";
}

// Remove wire fees from credit notes and internal modifications
if ($invoice_type != 'INVOICE') {
	$wire_fee = 0;
}

// Display wire fees
$active_cell = 'F' . $row;
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, $fee_title);

$active_cell = 'I' . $row;
$cell_value = $wire_fee;
$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getNumberFormat()
	->setFormatCode('#,##0.00');
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, $cell_value);


$row++; // Advance to the next row
// Display prepaid amount
$active_cell = 'F' . $row;
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, 'Pre-Paid:');

$active_cell = 'I' . $row;
$cell_value = $invoice->data->amount_paid;
$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getNumberFormat()
	->setFormatCode('#,##0.00');
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, $cell_value);


$row++; // Advance to the next row
// Display grand total
$active_cell = 'D' . $row;
SetBold($objPHPExcel, $active_cell);
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, 'Remaining Balance');


$active_cell = 'F' . $row;
SetBold($objPHPExcel, $active_cell);
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, 'Total:');

// Show the currency of the invoice
$active_cell = 'G' . $row;
$cell_value = 'USD';
SetHorizontalAlignment($objPHPExcel, $active_cell, 'CENTER');
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, $cell_value);


$active_cell = 'I' . $row;
SetBold($objPHPExcel, $active_cell);
$cell_value = '=I' . ($usd_subtotal_row) . '-I' . ($row - 3) . '+I' . ($row - 2) . '-I' . ($row - 1);
$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getNumberFormat()
	->setFormatCode('#,##0.00');
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, $cell_value);

$row++; // Advance to next row
// Surround all the data in a border
SetBorder($objPHPExcel, 'A11:I' . ($row), 'THIN');


$row++; // Advance to next row
// Spacer rows
$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(5);

$row++; // Advance to next row
// Display the cancellation policies headline
$active_cell = 'A' . $row;
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
SetBold($objPHPExcel, $active_cell);
SetHorizontalAlignment($objPHPExcel, $active_cell, 'CENTER');
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, 'Cancellation Policy');

$row++; // Advance to next row

$policy_counter = 1;

if ($invoice->data->use_custom_cx_policy) {
	$cancellation_policies = GetCustomCxPolicy($invoice);
}

// Show the individual policies
if (isset ($cancellation_policies)) {
	foreach ($cancellation_policies as $policy) {

		$active_cell = 'A' . $row;
		$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
		$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
		$cell_value = $policy;
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue($active_cell, $cell_value);

		$row++; // Advance to the next row

		$policy_counter++;
	}
}


// Surround all the policies in a border
SetBorder($objPHPExcel, ('A' . ($row - $policy_counter) . ':I' . ($row - 1)), 'THIN');

// Spacer rows
$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(5);

$row++; // Advance to the next row
//=========================Display the payment policy headline================================
$payment_policies = $invoice->GetPaymentPolicies();

$row_start_policy = $row;
$active_cell = 'A' . $row;
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
SetBold($objPHPExcel, $active_cell);
SetHorizontalAlignment($objPHPExcel, $active_cell, 'CENTER');
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, 'Payment Policy');

$row++; // Advance to next row
// Display the currency disclaimer if any currency other than USD is detected.
// If there is more than one element in the currencies array then we have USD + some other currency
if (count($currencies_array) > 1) {
	$currency_compare_str = 'The final balance in ';

	// Get all the currencies used
	foreach ($currencies_array as $currency_code => $value) {
		if ($currency_code != 'USD') {
			$currency_compare_str .= "USD vs $currency_code, ";
		}
	}

	$currency_compare_str .= 'will be reviewed when the balance is paid for.';

	// Print the first line
	$active_cell = 'A' . $row;
	SetBold($objPHPExcel, $active_cell);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'Note that the exchange is provided as an indication.');

	$row++;

	// Print the second line
	$active_cell = 'A' . $row;
	SetBold($objPHPExcel, $active_cell);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, $currency_compare_str);

	$row++;
}

foreach ($payment_policies as $policy) {

	if ($policy->due_date == '') {
		continue;
	}

	$due_date = date("M d-Y", strtotime($policy->due_date));
	$policy_text = "{$policy -> amount_due_percent}% Due upon receipt. Remaining balance due on $due_date.";
	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);

	SetHorizontalAlignment($objPHPExcel, $active_cell, 'LEFT');
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, "$policy_text");

	$row++; // Advance to next row
}

//$additionalPolicies = array(
//		"Any refund of credit card deposits due to cancellations, amendments or partial refund would be subject to a 3% non refundable processing fee. For payment by bank transfer, please be advised that you are responsible for the bank fees. A copy of the bank receipt is necessary for tracking purposes."
//	);
//foreach ($additionalPolicies as $policy_text) {
//	$active_cell = 'A' . $row;
//	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
//	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(9);
//	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getAlignment()->setWrapText(true);
//	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(40);
//	SetHorizontalAlignment($objPHPExcel, $active_cell, 'LEFT');
//	$objPHPExcel->setActiveSheetIndex(0)
//		->setCellValue($active_cell, "$policy_text");
//
//	$row++; // Advance to next row
//
//}
//
//// Surround the bank info in a border
SetBorder($objPHPExcel, ('A' . ($row_start_policy) . ':I' . ($row - 1)), 'THIN');

//=================================================================
// Spacer rows
$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(5);

$row++; // Advance to the next row


if($withAuthorizationForm){

// Display the bank information
	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(15);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(12);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setBold(true);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'Authorization of charge');

	$row++; // Advance to next row
	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(18);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, "Cardholder's Name: _____________________________________________________________________");


    $row++; // Advance to next row

//    $active_cell = 'A' . $row;
//    $objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(18);
//    $objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
//    $objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
//    $objPHPExcel->setActiveSheetIndex(0)
//        ->setCellValue($active_cell, 'Last 4# of CC: _____________________________________________________________________');
//
//
//    $row++; // Advance to next row

	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(18);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'Amount Authorized: _____________________________________________________________________');

	$row++; // Advance to next row

	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(36);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(9);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'I herewith authorize Overseas Travel of Florida, LLC. or its affiliated companies to charge the above-stated amount on the card provided and confirm that I have read and agree with the Cancellation and Payment policies indicated on this invoice');


	$row++; // Advance to next row

	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(18);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'Authorized signature: ____________________________________________________________________');

	$row++; // Advance to next row

	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(18);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'Date: _____________________________');

	$row++; // Advance to next row

	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(12);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'A form of ID corresponding to the credit card used is required along with this Authorization');

// Surround the bank info in a border
	SetBorder($objPHPExcel, ('A' . ($row - 6) . ':I' . ($row)), 'THIN');

	$objPHPExcel->getActiveSheet()->getStyle('A' . ($row - 6) . ':I' . ($row))->applyFromArray(
		array(
			'fill' => array(
				'type' => PHPExcel_Style_Fill::FILL_SOLID,
				'color' => array('rgb' => 'FFFF00')
			)
		)
	);
}else{
// Display the bank information
	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(12);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'Bank Details: Account holder name: Overseas Travel Of Florida, LLC.');

	$row++; // Advance to next row
	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(12);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, ' JP Morgan Chase Bank NA, 355 Alhambra Circle Floor 1, Coral Gables FL 33134');

	$row++; // Advance to next row

	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(12);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'Tel: 1.877.425.8100   Fax: 1.855.752.4344   Contact: Mr. James Pensado');

	$row++; // Advance to next row

	$active_cell = 'A' . $row;
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(12);
	$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
	$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue($active_cell, 'Account #: 398022605 - SWIFT for Intl wires: CHASUS33 - (Routing/ABA for US Wires only: 267084131)');

// Surround the bank info in a border
	SetBorder($objPHPExcel, ('A' . ($row - 3) . ':I' . ($row)), 'THIN');

	$objPHPExcel->getActiveSheet()->getStyle('A' . ($row - 3) . ':I' . ($row))->applyFromArray(
		array(
			'fill' => array(
				'type' => PHPExcel_Style_Fill::FILL_SOLID,
				'color' => array('rgb' => 'FFFF00')
			)
		)
	);
}



$row++; // Advance to next row
// Spacer rows
$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(5);

$row++; // Advance to the next row
// Display the Footer
$active_cell = 'A' . $row;
$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(12);
$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
SetHorizontalAlignment($objPHPExcel, $active_cell, 'CENTER');
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, 'OVERSEAS TRAVEL OF FLORIDA, LLC - SELLERS OF TRAVEL No: ST32773');

$row++; // Advance to next row

$active_cell = 'A' . $row;
$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(12);
$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
SetHorizontalAlignment($objPHPExcel, $active_cell, 'CENTER');
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, '814 Ponce de Leon Blvd Suite 400, Coral Gables, FL 33134, USA');

$row++; // Advance to next row

$active_cell = 'A' . $row;
$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(12);
$objPHPExcel->getActiveSheet()->getStyle($active_cell)->getFont()->setSize(10);
SetHorizontalAlignment($objPHPExcel, $active_cell, 'CENTER');
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':I' . $row);
$objPHPExcel->setActiveSheetIndex(0)
	->setCellValue($active_cell, 'Tel: 1-786-276-8686   -   Fax: 1-786-276-8646   -   Email: ar@overseasinternational.com');


// Surround the bank info in a border
SetBorder($objPHPExcel, ('A' . ($row - 2) . ':I' . ($row)), 'THIN');
$objPHPExcel->getActiveSheet()->getStyle('A' . ($row - 2) . ':I' . ($row))->applyFromArray(
	array(
		'fill' => array(
			'type' => PHPExcel_Style_Fill::FILL_SOLID,
			'color' => array('rgb' => 'FFFF00')
		)
	)
);

// Save Excel 2007 file
//$objWriter = PHPExcel_IOFactory::createWriter ( $objPHPExcel, 'Excel2007' );

$objWriter = new \PHPExcel_Writer_Excel2007($objPHPExcel);
//-- Calculate formulas
$objWriter->setPreCalculateFormulas();


//Define if it is either an invoice or a credit note
$inv_type = stripos($invoice_number, 'C') ? 'Credit_Note' : (stripos($invoice_number, 'INT') ? 'Int_Mod' : 'Inv');
$filename = __DIR__ . "/../../../invoices/$inv_type - " . $quote_data['client_name'] . ' (' . $agency_info['agency_name'] . ') - ' . ' (' . $invoice_number . ').xlsx';
$aux = basename($filename);

header("Pragma: public");
header("Expires: 0");
header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
header("Content-Type: application/force-download");
header("Content-Type: application/octet-stream");
header("Content-Type: application/download");
header("Content-Transfer-Encoding: binary ");
header("Content-Disposition: attachment; filename=\"$aux\"");

// If a request for an invoice preview is detected, save the file. Otherwise output to the buffer
if (isset ($_POST['invoice_preview'])) {
	$filename = __DIR__ . '/../../../invoices/' . $invoice_id . '.xlsx';
	$objWriter->save($filename);
} else {
	$objWriter->save('php://output');
}

//readfile($filename);

/**
 *
 * @param type $objPHPExcel
 * @param type $cell
 * @param type $border_style
 *
 */
function SetBorder(&$objPHPExcel, $cell, $border_style)
{

	switch ($border_style) {
		case 'THICK':
			$border_style = PHPExcel_Style_Border::BORDER_THICK;
			break;
		case 'MEDIUM':
			$border_style = PHPExcel_Style_Border::BORDER_MEDIUM;
			break;
		case 'THIN':
			$border_style = PHPExcel_Style_Border::BORDER_THIN;
			break;
	}


	$objPHPExcel->getActiveSheet()->getStyle($cell)
		->getBorders()->getTop()->setBorderStyle($border_style);
	$objPHPExcel->getActiveSheet()->getStyle($cell)
		->getBorders()->getBottom()->setBorderStyle($border_style);
	$objPHPExcel->getActiveSheet()->getStyle($cell)
		->getBorders()->getLeft()->setBorderStyle($border_style);
	$objPHPExcel->getActiveSheet()->getStyle($cell)
		->getBorders()->getRight()->setBorderStyle($border_style);
}

/**
 *
 * @param type $objPHPExcel
 * @param type $cell
 *
 */
function SetBold(&$objPHPExcel, $cell)
{
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getFont()->setBold(true);
}

/**
 *
 * @param type $objPHPExcel
 * @param type $cell
 * @param type $alignment
 *
 */
function SetHorizontalAlignment(&$objPHPExcel, $cell, $alignment)
{

	switch ($alignment) {
		case 'LEFT':
			$alignment_style = PHPExcel_Style_Alignment::HORIZONTAL_LEFT;
			break;
		case 'CENTER':
			$alignment_style = PHPExcel_Style_Alignment::HORIZONTAL_CENTER;
			break;
		case 'RIGHT':
			$alignment_style = PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;
			break;
	}

	$objPHPExcel->getActiveSheet()->getStyle($cell)
		->getAlignment()->setHorizontal($alignment_style);
}

function GetCustomCxPolicy($invoice)
{
	global $dashboard_pdo;
	$output = array();

	$cx_policy = new InvItemCxPolicy();

	$inv_check_in = $invoice->GetHighestCheckInDate();

	$query = "SELECT * FROM oti_custom_cx_policies WHERE invoice_id = :invoice_id";
	$stmt = $dashboard_pdo->prepare($query);
	$stmt->bindValue(':invoice_id', $invoice->data->invoice_id);

	$stmt->execute();

	while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
		$cx_policy->set_volatile(0, $row['cancel_days'], $row['cancel_time'], $row['cancel_fee'], $row['cancel_fee_type'], $row['cancel_code'], $inv_check_in, 'USD');
		$policy_description = $cx_policy->get_policy_description(0);

		switch ($row['item_category']) {
			case 'hotel':

				$output[] = 'Hotels: ' . $policy_description;
				break;
			case 'activity':
				$output[] = 'Activities: ' . $policy_description;
				break;
			case 'car_rental':
				$output[] = 'Car Rentals: ' . $policy_description;
				break;
			case 'transfer':
				$output[] = 'Transfers: ' . $policy_description;
				break;
		}
	}

	return $output;
}
