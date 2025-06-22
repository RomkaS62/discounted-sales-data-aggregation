<?php

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
/**
 * @package discounted_sales_data_aggregation
 * @version 1.0.0
 */
/*
Plugin Name: Aggregation of discounted sale data
Plugin URI: http://localhost:80
Description: Aggregates data concerning discounted sales once a day
Author: Ramūnas Mažeikis
Version: 1.0.0
Author URI: http://localhost:80
*/

// Do not load directly.
if (!defined('ABSPATH')) {
    die();
}

require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

const ISO_DATE_FORMAT = 'Y-m-d';
const ISO_TIMESTAMP_FORMAT = 'Y-m-d H:i:s';

add_action('woocommerce_init', 'dsda_init');

function dsda_init()
{
    add_action(
        'dsda_export_discounted_sales_data_hook',
        'dsda_export_discounted_sales_data'
    );

    if (dsda_debug()) {
        add_filter('cron_schedules', 'dsda_define_wp_cron_schedules');
    }

    $scheduled_time = wp_next_scheduled('dsda_export_discounted_sales_data_hook');
    $expected_execution_time = dsda_tomorrow_timestamp() + 600;

    if (dsda_force_run()) {
        wp_unschedule_event($scheduled_time, 'dsda_export_discounted_sales_data_hook');
        update_option('dsda_force', false);
        $expected_execution_time = time();
    }

    if ($scheduled_time != $expected_execution_time) {
        dsda_debug_log("Scheduling event for " . date(ISO_TIMESTAMP_FORMAT, $expected_execution_time));
        wp_schedule_event($expected_execution_time, 'daily', 'dsda_export_discounted_sales_data_hook');
    } else {
        $formatted_time = date(ISO_TIMESTAMP_FORMAT, $scheduled_time);
        dsda_debug_log("Event already scheduled for $formatted_time.");
    }
}

function dsda_define_wp_cron_schedules($schedules)
{
    $schedules['once_a_second'] = array(
        'interval' => 1,
        'display' => __('Once every second')
    );

    return $schedules;
}

function dsda_export_discounted_sales_data(): void
{
    dsda_debug_log("Running job");

    $date_completed = date(ISO_DATE_FORMAT, strtotime("-1 days"));

    if (dsda_debug()) {
        $date_completed = '1984-01-01...2077-01-01';
    }

    dsda_debug_log("Getting orders ($date_completed)");
    $orders = wc_get_orders([
        'date_completed' => $date_completed,
        'status' => 'wc-completed',
        'limit' => -1
    ]);

    $formatted_orders = [];
    $customers = [];
    $products = [];

    dsda_debug_log("Processing...");

    foreach ($orders as $order) {
        $user = $order->get_user();

        if ($user !== false) {
            $billing_email = $order->get_billing_email();

            if (!array_key_exists($billing_email, $customers)) {
                $customers[$billing_email] = [
                    'first_name' => $order->get_billing_first_name(),
                    'last_name' => $order->get_billing_last_name(),
                    'email' => $order->get_billing_email(),
                    'billing_phone' => $order->get_billing_phone()
                ];
            }
        }

        foreach ($order->get_items() as $item) {
            $product = $item->get_product();

            $regular_subtotal = $product->get_regular_price() * $item->get_quantity();
            $factual_subtotal = $item->get_subtotal();
            $sold_at_a_discount = $factual_subtotal < $regular_subtotal;

            $formatted_orders[] = [
                'date'                  => $order->get_date_created(),
                'order_number'          => $order->get_order_number(),
                'customer'              => $order->get_customer_id(),
                'item_name'             => $item->get_name(),
                'sold_at_a_discount'    => $sold_at_a_discount ? 'yes' : 'no',
                'quantity'              => $item->get_quantity(),
                'sum'                   => $item->get_subtotal()
            ];

            if (!array_key_exists($product->get_id(), $products)) {
                $products[$product->get_id()] = [
                    'id'                            => $product->get_id(),
                    'name'                          => $product->get_name(),
                    'remainder'                     => $product->get_stock_quantity(),
                    'number_sold'                   => 0,
                    'number_sold_under_discount'    => 0
                ];
            }

            $products[$product->get_id()]['number_sold'] += $item->get_quantity();

            if ($sold_at_a_discount) {
                $products[$product->get_id()]['number_sold_under_discount'] += $item->get_quantity();
            }
        }
    }

    dsda_debug_log("Creating spreadsheet");
    $spreadsheet = new Spreadsheet();

    $customer_sheet = $spreadsheet->getActiveSheet();
    $customer_sheet->setTitle("Customers");

    $order_sheet = $spreadsheet->createSheet(1);
    $order_sheet->setTitle("Orders");

    $product_sheet = $spreadsheet->createSheet(2);
    $product_sheet->setTitle("Products");

    dsda_debug_log("Writing data to spreadsheet");

    dsda_to_xlsx($customer_sheet, array_values($customers), [
        ['key' => 'first_name',                 'name' => 'First name'],
        ['key' => 'last_name',                  'name' => 'Last name'],
        ['key' => 'email',                      'name' => 'email'],
        ['key' => 'billing_phone',              'name' => 'Billing phone']
    ]);

    dsda_to_xlsx($order_sheet, array_values($formatted_orders), [
        ['key' => 'date',                       'name' => 'Date completed'],
        ['key' => 'order_number',               'name' => 'Order number'],
        ['key' => 'customer',                   'name' => 'Customer ID'],
        ['key' => 'item_name',                  'name' => 'Item name'],
        ['key' => 'sold_at_a_discount',         'name' => 'Sold at a discount'],
        ['key' => 'quantity',                   'name' => 'Quantity sold'],
        ['key' => 'sum',                        'name' => 'Toal value']
    ]);

    dsda_to_xlsx($product_sheet, array_values($products), [
        ['key' => 'id',                         'name' => 'ID'],
        ['key' => 'name',                       'name' => 'Name'],
        ['key' => 'remainder',                  'name' => 'Remainder'],
        ['key' => 'number_sold',                'name' => 'Number sold'],
        ['key' => 'number_sold_under_discount', 'name' => 'Number sold under discount']
    ]);

    $output_dir = get_option('dsda_output_dir');
    $output_path = ABSPATH . "$output_dir/$date_completed.xlsx";
    dsda_debug_log("Writing data to file: $output_path");

    $writer = new Xlsx($spreadsheet);
    $writer->save($output_path);
}

function dsda_to_xlsx(Worksheet $sheet, array $data, array $spec): void
{
    for ($i = 0; $i < count($spec); $i++) {
        $col_display_name = $spec[$i]['name'];
        $sheet->setCellValue([$i + 1, 1], $col_display_name);
    }

    for ($i = 0; $i < count($data); $i++) {
        $row = $data[$i];

        for ($j = 0; $j < count($spec); $j++) {
            $key = $spec[$j]['key'];
            $value = $row[$key];

            if ($value) {
                $sheet->setCellValue([$j + 1, $i + 2], $value);
            }
        }
    }
}

function dsda_debug_log(string $str): void
{
    if (dsda_debug())
        error_log($str);
}

function dsda_debug(): bool
{
    return get_option('dsda_debug');
}

function dsda_force_run(): bool
{
    return get_option('dsda_force');
}

add_action('admin_menu', function() {
    add_options_page('DSDA Settings', 'DSDA Settings', 'manage_options', 'dsda-settings', 'dsda_settings_page');
});

add_action('admin_init', function() {
    register_setting('dsda_settings_group', 'dsda_debug');
    register_setting('dsda_settings_group', 'dsda_force');
    register_setting('dsda_settings_group', 'dsda_output_dir');

    add_settings_section('dsda_section', 'General', null, 'dsda-settings');

    dsda_add_bool_settings_field('dsda_debug', 'Debug mode');
    dsda_add_bool_settings_field('dsda_force', 'Force execution');
    dsda_add_string_settings_field('dsda_output_dir', 'Path to output directory');
});

function dsda_add_bool_settings_field(string $name, string $description): void
{
    add_settings_field($name, $description, function() use ($name) {
        $opt_value = get_option($name);
        echo '<input type="checkbox" name="' . $name . '" value="1" ' . dsda_checked($opt_value) .'/>';
    }, 'dsda-settings', 'dsda_section');
}

function dsda_add_string_settings_field(string $name, string $description): void
{
    add_settings_field($name, $description, function() use ($name) {
        $opt_value = get_option($name);
        echo '<input type="text" name="' . $name . '" value="' . esc_attr($opt_value) . '"/>';
    }, 'dsda-settings', 'dsda_section');
}

function dsda_checked(bool $val): string
{
    if ($val)
        return 'checked';

    return '';
}

function dsda_settings_page() {
    ?>
    <div class="wrap">
        <h1>DSDA Settings</h1>
        <form method="post" action="options.php">
            <?php
            settings_fields('dsda_settings_group');
            do_settings_sections('dsda-settings');
            submit_button();
            ?>
        </form>
    </div>
    <?php
}

function dsda_tomorrow_timestamp()
{
    $DAY_IN_SECONDS = 24 * 60 * 60;
    return floor(strtotime("+1 days") / $DAY_IN_SECONDS) * $DAY_IN_SECONDS;
}
