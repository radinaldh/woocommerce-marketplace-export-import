<?php

/**
 * Plugin Name: WooCommerce Marketplace Export/Import
 * Description: A custom plugin for exporting products in Tokopedia format.
 * Version: 1.5.0
 * Author: Radinal
 * License: GPL2
 */

if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
}

if (!defined('ABSPATH')) {
    exit;
}

// Import necessary classes
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Add a new menu called "Marketplace" with a submenu "Tokopedia Export"
add_action('admin_menu', 'add_marketplace_menu');

function add_marketplace_menu()
{
    add_menu_page(
        'Marketplace',              // Page title
        'Marketplace',              // Menu title
        'manage_options',           // Capability required to view the menu
        'marketplace',              // Menu slug
        'marketplace_overview_page', // Function to display the page content
        'dashicons-store',          // Icon for the menu
        58                          // Position in the menu
    );

    add_submenu_page(
        'marketplace',              // Parent slug
        'Tokopedia Export',         // Page title
        'Tokopedia Export',         // Menu title
        'manage_options',           // Capability required to view the menu
        'tokopedia-export',         // Menu slug
        'tokopedia_export_page'     // Function to display the page content
    );

    add_submenu_page(
        'marketplace',              // Parent slug
        'Blibli Export',            // Page title
        'Blibli Export',            // Menu title
        'manage_options',           // Capability required to view the menu
        'blibli-export',            // Menu slug
        'blibli_export_page'        // Function to display the page content
    );
}

// Display the Marketplace overview page
function marketplace_overview_page()
{
?>
    <div class="wrap">
        <h1>Marketplace Overview</h1>
        <p>Welcome to the Marketplace management page. Here you can manage exports and imports for Tokopedia and Blibli.</p>
        <ul>
            <li><a href="<?php echo admin_url('admin.php?page=tokopedia-export'); ?>">Tokopedia Export</a></li>
            <li><a href="<?php echo admin_url('admin.php?page=blibli-export'); ?>">Blibli Export</a></li>
            <li><a href="<?php echo admin_url('admin.php?page=tokopedia-import'); ?>">Tokopedia Import</a></li>
            <li><a href="<?php echo admin_url('admin.php?page=blibli-import'); ?>">Blibli Import</a></li>
        </ul>
    </div>
<?php
}


function enqueue_select2_assets()
{
    // Enqueue Select2 CSS
    wp_enqueue_style('select2-css', 'https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css');

    // Enqueue Select2 JS
    wp_enqueue_script('select2-js', 'https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js', array('jquery'), null, true);

    // Enqueue your custom Select2 initialization script
    wp_enqueue_script('custom-select2-init', plugins_url('/js/select2-init.js', __FILE__), array('select2-js'), null, true);
}
add_action('admin_enqueue_scripts', 'enqueue_select2_assets');

// Display the Tokopedia Export page
function tokopedia_export_page()
{
?>
    <div class="wrap">
        <h1>Tokopedia Export</h1>
        <form method="post" action="<?php echo admin_url('admin-post.php'); ?>">
            <input type="hidden" name="action" value="export_tokopedia_xlsx">
            <table class="form-table">
                <tr valign="top">
                    <th scope="row">Product Category</th>
                    <td>
                        <?php
                        // Get product categories
                        $args = array(
                            'taxonomy'   => 'product_cat',
                            'hide_empty' => false,
                        );
                        $categories = get_terms($args);

                        // Dropdown for product categories with Select2 integration and multiple select enabled
                        echo '<select name="export_category[]" multiple="multiple" style="width: 100%;">';
                        foreach ($categories as $category) {
                            echo '<option value="' . esc_attr($category->term_id) . '">' . esc_html($category->name) . '</option>';
                        }
                        echo '</select>';
                        ?>
                    </td>
                </tr>
                <tr valign="top">
                    <th scope="row">Select Products</th>
                    <td>
                        <?php
                        // Get all products
                        $products = get_posts(array(
                            'post_type' => 'product',
                            'posts_per_page' => -1,
                            'orderby' => 'title',
                            'order' => 'ASC',
                        ));

                        // Dropdown for products with Select2 integration and multiple select enabled
                        echo '<select name="export_products[]" multiple="multiple" style="width: 100%;">';
                        foreach ($products as $product) {
                            echo '<option value="' . esc_attr($product->ID) . '">' . esc_html($product->post_title) . '</option>';
                        }
                        echo '</select>';
                        ?>
                    </td>
                </tr>
                <tr valign="top">
                    <th scope="row">Exclude Draft & Private Products</th>
                    <td>
                        <input type="checkbox" name="exclude_draft_private" value="1">
                    </td>
                </tr>
            </table>
            <p class="submit">
                <input type="submit" name="submit" class="button-primary" value="Generate Tokopedia XLSX">
            </p>
        </form>
    </div>
<?php
}

add_action('admin_post_export_tokopedia_xlsx', 'generate_tokopedia_xlsx');
function generate_tokopedia_xlsx()
{
    // Path to the template file
    $template_path = __DIR__ . '/sample_tokped.xlsx';

    // Load the template spreadsheet
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($template_path);
    $sheet = $spreadsheet->getActiveSheet();

    // Fetch the selected product categories and products
    $categories = isset($_POST['export_category']) ? array_map('intval', $_POST['export_category']) : [];
    $selected_products = isset($_POST['export_products']) ? array_map('intval', $_POST['export_products']) : [];
    $exclude_draft_private = isset($_POST['exclude_draft_private']) ? true : false;

    // Build the query arguments based on user input
    $args = [
        'post_type'      => 'product',
        'posts_per_page' => -1,
        'orderby'        => 'date',
        'order'          => 'DESC',
    ];

    if ($exclude_draft_private) {
        $args['post_status'] = ['publish']; // Exclude draft and private products
    }

    if (!empty($categories)) {
        $args['tax_query'] = [
            [
                'taxonomy' => 'product_cat',
                'field'    => 'term_id',
                'terms'    => $categories,
            ],
        ];
    }

    if (!empty($selected_products)) {
        $args['post__in'] = $selected_products;
    }

    $products = new WP_Query($args);

    // Define the mapping for brands to Column G values
    $brand_mapping = [
        'BALENCIAGA' => '17388550',
        'BOTTEGA VENETA' => '17388579',
        'CHANEL' => '17388526',
        'CHLOÉ' => '17388556',
        'CELINE' => '17388628',
        'DELVAUX' => '17388590',
        'FAURÉ LE PAGE' => '17388529',
        'FENDI' => '17388562',
        'GIVENCHY' => '17388596',
        'GUCCI' => '17388635',
        'HERMÈS' => '17513205',
        'LOUIS VUITTON' => '17388565',
        'MIU MIU' => '17388532',
        'MARC JACOB' => '17388606',
        'MCM' => '17388637',
        'PHILIP LIM' => '17388570',
        'PROENZA SCHOULER' => '17388645',
        'PRADA' => '17388613',
        'SAINT LAURENT' => '17388547',
        'SALVATORE FERRAGAMO' => '17388574',
        'TOD\'S' => '17388618',
        'VALENTINO' => '17388652',
        'BVLGARI' => '18416049',
        'CHOPARD' => '19406932',
        'GOYARD' => '17388630',
        'TOD' => '20650613',
        'ROLEX' => '22705985',
        'DIOR' => '25220080',
        'LOEWE' => '27366907',
        'SECOND CHANCE LIVE' => '29839735',
        'ALAÏA' => '29890594',
        'DE LA COUR' => '30446513',
        'BAO BAO' => '30950622',
    ];

    if ($products->have_posts()) {
        $row_index = 4; // Start adding products from the fourth row
        while ($products->have_posts()) {
            $products->the_post();
            global $product;

            // Check if ACF field button_tokopedia is empty
            $tokopedia_button = get_field('button_tokopedia', $product->get_id());
            if (!empty($tokopedia_button)) {
                continue; // Skip if the button_tokopedia field is not empty
            }

            // Fetch the brand from the product's meta data
            $brand = get_post_meta($product->get_id(), '_brand', true);
            $product_name = $product->get_name();
            $sku = $product->get_sku();

            // Format the output for Column B
            $formatted_name = sprintf('%s %s %s', $brand, $product_name, $sku);

            // Determine the value for Column G based on the brand
            $column_g_value = isset($brand_mapping[strtoupper($brand)]) ? $brand_mapping[strtoupper($brand)] : ''; // Default to nothing

            // Prepare the custom text to be added above the product description in Column C
            $custom_text_top = "REMINDER: Teliti sebelum membeli, untuk foto & keterangan lebih lengkap silahkan hubungi kami melalui chat\n\n";
            $custom_text_bottom = "\n\nMore Detail Please Visit Our Website - https://secondchancebag.com\n\n" . "Note :\n\n" . "Due to the nature of online sales, the color of the image photograph and the actual product may differ slightly depending on the monitor environment of the personal computer or smartphone that is being used by customers, shooting, image quality and so on.\n\n";

            // Combine the custom text with the product description
            $description = $custom_text_top . $product->get_description() . $custom_text_bottom;

            $product_data = [
                '', // Column A
                $formatted_name, // Column B
                $description, // Column C
                '1919', // Column D
                $product->get_weight(), // Column E
                '1', // Column F
                $column_g_value, // Column G
                '', // Column H
                'Bekas', // Column I
                wp_get_attachment_url($product->get_image_id()), // Column J (featured image URL)
            ];

            // Get the product gallery images
            $gallery_image_ids = $product->get_gallery_image_ids();

            // Fill columns K, L, M, N with gallery images
            for ($i = 0; $i < 4; $i++) {
                $product_data[] = isset($gallery_image_ids[$i]) ? wp_get_attachment_url($gallery_image_ids[$i]) : '';
            }

            // Calculate the price with a 7% increase and round it to the nearest 100,000
            $regular_price = $product->get_regular_price();
            $price_with_markup = $regular_price / 0.92;
            $rounded_price = ceil($price_with_markup / 100000) * 100000;

            // Add the rest of the columns
            $product_data = array_merge($product_data, [
                $product->get_sku(), // Column O
                'Nonaktif', // Column P
                '1', // Column Q
                $rounded_price, // Column R (Regular price + 7% rounded to nearest 100,000)
                '', // Column S
                'opsional', // Column T
            ]);

            foreach ($product_data as $col_index => $value) {
                $sheet->setCellValueByColumnAndRow($col_index + 1, $row_index, $value);
            }

            $row_index++;
        }

        // Save the file
        $file_name = 'tokopedia_export_' . date('Y-m-d_H-i-s') . '.xlsx';
        $file_path = wp_upload_dir()['path'] . '/' . $file_name;

        $writer = new Xlsx($spreadsheet);
        $writer->save($file_path);

        // Serve the file for download
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . basename($file_path) . '"');
        header('Content-Length: ' . filesize($file_path));
        readfile($file_path);
        exit;
    } else {
        wp_redirect(add_query_arg(['error' => 'no_products'], wp_get_referer()));
        exit;
    }

    wp_reset_postdata();
}

function blibli_export_page()
{
?>
    <div class="wrap">
        <h1>Blibli Export</h1>
        <form method="post" action="<?php echo admin_url('admin-post.php'); ?>">
            <table class="form-table">
                <tr valign="top">
                    <th scope="row">Product Category</th>
                    <td>
                        <?php
                        // Get product categories
                        $args = array(
                            'taxonomy'   => 'product_cat',
                            'hide_empty' => false,
                        );
                        $categories = get_terms($args);

                        // Dropdown for product categories with Select2 integration and multiple select enabled
                        echo '<select name="export_category[]" multiple="multiple" style="width: 100%;">';
                        foreach ($categories as $category) {
                            echo '<option value="' . esc_attr($category->term_id) . '">' . esc_html($category->name) . '</option>';
                        }
                        echo '</select>';
                        ?>
                    </td>
                </tr>
                <tr valign="top">
                    <th scope="row">Select Products</th>
                    <td>
                        <?php
                        // Get all products
                        $products = get_posts(array(
                            'post_type' => 'product',
                            'posts_per_page' => -1,
                            'orderby' => 'title',
                            'order' => 'ASC',
                        ));

                        // Dropdown for products with Select2 integration and multiple select enabled
                        echo '<select name="export_products[]" multiple="multiple" style="width: 100%;">';
                        foreach ($products as $product) {
                            echo '<option value="' . esc_attr($product->ID) . '">' . esc_html($product->post_title) . '</option>';
                        }
                        echo '</select>';
                        ?>
                    </td>
                </tr>
                <tr valign="top">
                    <th scope="row">Exclude Draft & Private Products</th>
                    <td>
                        <input type="checkbox" name="exclude_draft_private" value="1">
                    </td>
                </tr>
            </table>
            <p class="submit">
                <input type="submit" name="action" value="export_blibli_xlsx" class="button-primary" style="margin-right: 10px;">
                <input type="submit" name="action" value="export_blibli_xlsx_sale" class="button-primary">
            </p>
        </form>
    </div>
<?php
}

function generate_blibli_xlsx()
{
    // Path to the template file
    $template_path = __DIR__ . '/sample_blibli.xlsx';

    // Load the template spreadsheet
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($template_path);
    $sheet = $spreadsheet->getActiveSheet();

    // Fetch the selected product categories and products
    $categories = isset($_POST['export_category']) ? array_map('intval', $_POST['export_category']) : [];
    $selected_products = isset($_POST['export_products']) ? array_map('intval', $_POST['export_products']) : [];
    $exclude_draft_private = isset($_POST['exclude_draft_private']) ? true : false;

    // Build the query arguments based on user input
    $args = [
        'post_type'      => 'product',
        'posts_per_page' => -1,
        'orderby'        => 'date',
        'order'          => 'DESC',
    ];

    if ($exclude_draft_private) {
        $args['post_status'] = ['publish']; // Exclude draft and private products
    }

    if (!empty($categories)) {
        $args['tax_query'] = [
            [
                'taxonomy' => 'product_cat',
                'field'    => 'term_id',
                'terms'    => $categories,
            ],
        ];
    }

    if (!empty($selected_products)) {
        $args['post__in'] = $selected_products;
    }

    $products = new WP_Query($args);

    if ($products->have_posts()) {
        $row_index = 5; // Start adding products from the fifth row
        while ($products->have_posts()) {
            $products->the_post();
            global $product;

            // Check if ACF field button_blibli is empty
            $blibli_button = get_field('button_blibli', $product->get_id());
            if (!empty($blibli_button)) {
                continue; // Skip if the button_blibli field is not empty
            }

            // Fetch the brand from the product's meta data
            $brand = get_post_meta($product->get_id(), '_brand', true);
            $product_name = $product->get_name();
            $sku = $product->get_sku();

            // Format the output for Column B
            $formatted_name = sprintf('%s %s %s', $brand, $product_name, $sku);

            $description = $product->get_description();

            // Convert newlines to <br> tags
            $description = nl2br(trim($description));

            // Optionally, remove extra spaces if any were introduced
            $description = preg_replace('/\s+/', ' ', $description);

            // Prepare the custom text to be added above and below the product description
            $custom_text_top = "REMINDER: Teliti sebelum membeli, untuk foto & keterangan lebih lengkap silahkan hubungi kami melalui chat<br><br>";
            $custom_text_bottom = "Note :<br><br>" . "Due to the nature of online sales, the color of the image photograph and the actual product may differ slightly depending on the monitor environment of the personal computer or smartphone that is being used by customers, shooting, image quality and so on.<br><br>";

            // Combine the custom text with the product description
            $final_desc = $custom_text_top . $description . $custom_text_bottom;

            $regular_price = (float) $product->get_regular_price();
            $price_with_markup = round($regular_price / 0.97);
            $rounded_price = ceil($price_with_markup / 100000) * 100000;

            $sale_price = (float) $product->get_sale_price();
            $price_with_markup_sale = round($sale_price * 1.03);
            $rounded_price_sale = ceil($price_with_markup_sale / 100000) * 100000;

            $sale_start_date = get_post_meta($product->get_id(), '_sale_price_dates_from', true);
            $sale_end_date = get_post_meta($product->get_id(), '_sale_price_dates_to', true);

            if (!empty($sale_price) && empty($sale_start_date) && empty($sale_end_date)) {
                // Apply 3% markup to the sale price and round up
                $price_for_column_aa = $rounded_price_sale;
            } else {
                $price_for_column_aa = $rounded_price;
            }

            // Prepare product data array based on the Blibli format
            $product_data = [
                $formatted_name, // Column A: Formatted Name
                '', // Column B: Blank
                $sku, // Column C: SKU
                $final_desc, // Column D: Description with <br/> for line breaks
                '', // Column E: Blank
                'SECOND CHANCE', // Column F: Static Value "SECOND CHANCE"
                '', // Column G: Blank
                '', // Column H: Blank
                '', // Column I: Blank
                '', // Column J: Blank
                '', // Column K: Blank
                wp_get_attachment_url($product->get_image_id()), // Column L: Featured Image URL
            ];

            // Get the product gallery images
            $gallery_image_ids = $product->get_gallery_image_ids();

            // Fill columns M - R with gallery images
            for ($i = 0; $i < 6; $i++) {
                $product_data[] = isset($gallery_image_ids[$i]) ? wp_get_attachment_url($gallery_image_ids[$i]) : '';
            }

            // Continue with the remaining columns, including the calculated price for Column AA
            $product_data = array_merge($product_data, [
                '', // Column S: Blank
                'Melalui partner logistik Blibli', // Column T: Static Value "Blibli"
                'PP-3091873 || Second Chance Official Store', // Column U: Static Value "PP-3091873 || Second Chance Official Store"
                $product->get_length(), // Column V: Product Length
                $product->get_width(), // Column W: Product Width
                $product->get_height(), // Column X: Product Height
                $product->get_weight(), // Column Y: Product Weight
                $rounded_price, // Column Z: Price + 3% rounded
                $price_for_column_aa, // Column AA: Sale Price + 3% rounded up or Regular Price + 3% rounded
                '1', // Column AB: Static Value "1"
                '1', // Column AC: Static Value "1"
                '0', // Column AD: Static Value "0"
                '', // Column AE: Blank
                '', // Column AF: Blank
                '', // Column AG: Blank
                '', // Column AH: Blank
                '', // Column AI: Blank
                '', // Column AJ: Blank
                '', // Column AK: Blank
            ]);

            // Insert data into the spreadsheet
            foreach ($product_data as $col_index => $value) {
                $sheet->setCellValueByColumnAndRow($col_index + 1, $row_index, $value);
            }

            $row_index++;
        }

        // Save the file
        $file_name = 'blibli_export_' . date('Y-m-d_H-i-s') . '.xlsx';
        $file_path = wp_upload_dir()['path'] . '/' . $file_name;

        $writer = new Xlsx($spreadsheet);
        $writer->save($file_path);

        // Serve the file for download
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . basename($file_path) . '"');
        header('Content-Length: ' . filesize($file_path));
        readfile($file_path);
        exit;
    } else {
        wp_redirect(add_query_arg(['error' => 'no_products'], wp_get_referer()));
        exit;
    }

    wp_reset_postdata();
}
add_action('admin_post_export_blibli_xlsx', 'generate_blibli_xlsx');


function generate_blibli_xlsx_sale()
{
    // Path to the template file
    $template_path = __DIR__ . '/sample_sale_blibli.xlsx';

    // Load the template spreadsheet
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($template_path);
    $sheet = $spreadsheet->getActiveSheet();

    // Fetch the selected product categories and products
    $categories = isset($_POST['export_category']) ? array_map('intval', $_POST['export_category']) : [];
    $selected_products = isset($_POST['export_products']) ? array_map('intval', $_POST['export_products']) : [];

    // Build the query arguments based on user input
    $args = [
        'post_type'      => 'product',
        'posts_per_page' => -1,
        'orderby'        => 'date',
        'order'          => 'DESC',
    ];

    if (!empty($categories)) {
        $args['tax_query'] = [
            [
                'taxonomy' => 'product_cat',
                'field'    => 'term_id',
                'terms'    => $categories,
            ],
        ];
    }

    if (!empty($selected_products)) {
        $args['post__in'] = $selected_products;
    }

    $products = new WP_Query($args);

    if ($products->have_posts()) {
        $row_index = 12; // Start adding products from the twelfth row
        while ($products->have_posts()) {
            $products->the_post();
            global $product;

            // Fetch the brand and SKU from the product's meta data
            $brand = get_post_meta($product->get_id(), '_brand', true);
            $product_name = $product->get_name();
            $sku = $product->get_sku();

            // Format the output for Column C
            $formatted_name = sprintf('%s %s %s', $brand, $product_name, $sku);

            $regular_price = (float) $product->get_regular_price();
            $price_with_markup = round($regular_price / 0.97);
            $rounded_price = ceil($price_with_markup / 100000) * 100000;

            $sale_price = (float) $product->get_sale_price();
            $price_with_markup_sale = round($sale_price / 0.97);
            $rounded_price_sale = ceil($price_with_markup_sale / 100000) * 100000;


            // Check if the sale price has a time limit
            $sale_start_date = get_post_meta($product->get_id(), '_sale_price_dates_from', true);
            $sale_end_date = get_post_meta($product->get_id(), '_sale_price_dates_to', true);

            // Determine the price for Column E
            if (!empty($sale_price) && empty($sale_start_date) && empty($sale_end_date)) {
                $price_for_column_e = $rounded_price_sale;
            } else {
                $price_for_column_e = $rounded_price;
            }

            // Prepare product data array based on the Blibli sale format
            $product_data = [
                $sku, // Column A: SKU
                '', // Column B: Blank
                $formatted_name, // Column C: Formatted Name
                $rounded_price, // Column D: Regular Price
                $price_for_column_e, // Column E: Sale Price or Regular Price
                '1', // Column F: Static Value "1"
                '', // Column G: Blank
                '', // Column H: Blank
            ];

            // Insert data into the spreadsheet
            foreach ($product_data as $col_index => $value) {
                $sheet->setCellValueByColumnAndRow($col_index + 1, $row_index, $value);
            }

            $row_index++;
        }

        // Save the file
        $file_name = 'blibli_export_sale_' . date('Y-m-d_H-i-s') . '.xlsx';
        $file_path = wp_upload_dir()['path'] . '/' . $file_name;

        $writer = new Xlsx($spreadsheet);
        $writer->save($file_path);

        // Serve the file for download
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . basename($file_path) . '"');
        header('Content-Length: ' . filesize($file_path));
        readfile($file_path);
        exit;
    } else {
        wp_redirect(add_query_arg(['error' => 'no_products'], wp_get_referer()));
        exit;
    }

    wp_reset_postdata();
}
add_action('admin_post_export_blibli_xlsx_sale', 'generate_blibli_xlsx_sale');

/**
 * Import tokopedia
 */

// Add a new submenu for Tokopedia Import
add_action('admin_menu', 'add_tokopedia_import_menu');

function add_tokopedia_import_menu()
{
    add_submenu_page(
        'marketplace',               // Parent slug
        'Tokopedia Import',          // Page title
        'Tokopedia Import',          // Menu title
        'manage_options',            // Capability required to view the menu
        'tokopedia-import',          // Menu slug
        'tokopedia_import_page'      // Function to display the page content
    );
}

// Display the Tokopedia Import page
function tokopedia_import_page()
{
?>
    <div class="wrap">
        <h1>Tokopedia Import</h1>
        <form method="post" enctype="multipart/form-data" action="<?php echo admin_url('admin-post.php'); ?>">
            <input type="hidden" name="action" value="import_tokopedia_xlsx">
            <table class="form-table">
                <tr valign="top">
                    <th scope="row">Upload Tokopedia XLSX File</th>
                    <td>
                        <input type="file" name="tokopedia_xlsx" accept=".xlsx" required />
                    </td>
                </tr>
            </table>
            <p class="submit">
                <input type="submit" name="submit" class="button-primary" value="Import Tokopedia XLSX">
            </p>
        </form>
    </div>
<?php
}

// Handle the Tokopedia XLSX file import
add_action('admin_post_import_tokopedia_xlsx', 'handle_tokopedia_import');

function handle_tokopedia_import()
{
    if (isset($_FILES['tokopedia_xlsx']) && $_FILES['tokopedia_xlsx']['error'] == UPLOAD_ERR_OK) {
        // Load the uploaded file
        $file_path = $_FILES['tokopedia_xlsx']['tmp_name'];

        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file_path);
            $sheet = $spreadsheet->getActiveSheet();

            // Loop through the rows in the spreadsheet, starting from the fourth row (assuming the first row is headers)
            $highestRow = $sheet->getHighestRow();

            for ($row = 4; $row <= $highestRow; $row++) {
                // Get data from each column
                $formatted_name = trim($sheet->getCell('C' . $row)->getValue());
                $tokopedia_button = trim($sheet->getCell('D' . $row)->getValue());
                $stock = trim($sheet->getCell('F' . $row)->getValue());
                $tokopedia_price = trim($sheet->getCell('H' . $row)->getValue());
                $sku = trim($sheet->getCell('K' . $row)->getValue());
                $weight = trim($sheet->getCell('M' . $row)->getValue());

                // Use the SKU to check if the product exists
                if (empty($sku)) {
                    error_log("Skipped row $row: SKU is empty");
                    continue;
                }

                $product_id = wc_get_product_id_by_sku($sku);

                if ($product_id) {
                    // Update the existing product
                    $product = wc_get_product($product_id);
                } else {
                    // If the SKU does not exist, skip the product
                    error_log("Skipped row $row: SKU $sku does not exist");
                    continue;
                }

                $stock = is_numeric($stock) ? (int)$stock : 0;

                if ($stock > 0) {
                    $product->set_stock_quantity($stock);
                }

                $product->set_weight($weight);

                // Set Tokopedia price and ACF field button_tokopedia
                update_post_meta($product->get_id(), '_tokopedia_price', $tokopedia_price);
                update_field('button_tokopedia', $tokopedia_button, $product->get_id());

                // Save the product
                try {
                    $product_id = $product->save();
                    error_log("Product saved: $formatted_name (SKU: $sku, ID: $product_id)");
                    wp_cache_flush();
                } catch (Exception $e) {
                    error_log("Failed to save product: " . $e->getMessage());
                }
            }

            // Redirect to the import page with a success message
            wp_redirect(add_query_arg(['import_status' => 'success', 'import_source' => 'tokopedia'], admin_url('admin.php?page=tokopedia-import')));
            exit;
        } catch (Exception $e) {
            error_log("Error during import: " . $e->getMessage());
            wp_redirect(add_query_arg('import_status', 'error', admin_url('admin.php?page=tokopedia-import')));
            exit;
        }
    } else {
        wp_redirect(add_query_arg('import_status', 'error', admin_url('admin.php?page=tokopedia-import')));
        exit;
    }
}



function upload_image_from_url($url)
{
    $upload_dir = wp_upload_dir();
    $image_data = @file_get_contents($url, false, stream_context_create(['http' => ['timeout' => 10]]));

    if ($image_data) {
        $filename = basename($url);
        $file_path = $upload_dir['path'] . '/' . $filename;

        file_put_contents($file_path, $image_data);

        $attachment = array(
            'post_mime_type' => mime_content_type($file_path),
            'post_title'     => sanitize_file_name($filename),
            'post_content'   => '',
            'post_status'    => 'inherit'
        );

        $attach_id = wp_insert_attachment($attachment, $file_path);

        require_once(ABSPATH . 'wp-admin/includes/image.php');
        $attach_data = wp_generate_attachment_metadata($attach_id, $file_path);
        wp_update_attachment_metadata($attach_id, $attach_data);

        return $attach_id;
    } else {
        error_log("Failed to fetch image: $url");
        return false;
    }
}

add_action('admin_notices', 'tokopedia_import_notices');

function tokopedia_import_notices()
{
    if (isset($_GET['import_status']) && isset($_GET['import_source']) && $_GET['import_source'] == 'tokopedia') {
        if ($_GET['import_status'] == 'success') {
            echo '<div class="notice notice-success is-dismissible"><p>Tokopedia products have been successfully imported.</p></div>';
        } elseif ($_GET['import_status'] == 'error') {
            echo '<div class="notice notice-error is-dismissible"><p>There was an error importing the Tokopedia products. Please try again.</p></div>';
        }
    }
}

/**
 * Import Blibli
 */

// Add a new submenu for Blibli Import
add_action('admin_menu', 'add_blibli_import_menu');

function add_blibli_import_menu()
{
    add_submenu_page(
        'marketplace',               // Parent slug
        'Blibli Import',             // Page title
        'Blibli Import',             // Menu title
        'manage_options',            // Capability required to view the menu
        'blibli-import',             // Menu slug
        'blibli_import_page'         // Function to display the page content
    );
}

// Display the Blibli Import page
function blibli_import_page()
{
?>
    <div class="wrap">
        <h1>Blibli Import</h1>
        <form method="post" enctype="multipart/form-data" action="<?php echo admin_url('admin-post.php'); ?>">
            <input type="hidden" name="action" value="import_blibli_xlsx">
            <table class="form-table">
                <tr valign="top">
                    <th scope="row">Upload Blibli XLSX File</th>
                    <td>
                        <input type="file" name="blibli_xlsx" accept=".xlsx" required />
                    </td>
                </tr>
            </table>
            <p class="submit">
                <input type="submit" name="submit" class="button-primary" value="Import Blibli XLSX">
            </p>
        </form>
    </div>
<?php
}

// Handle the Blibli XLSX file import
add_action('admin_post_import_blibli_xlsx', 'handle_blibli_import');

function handle_blibli_import()
{
    if (isset($_FILES['blibli_xlsx']) && $_FILES['blibli_xlsx']['error'] == UPLOAD_ERR_OK) {
        // Load the uploaded file
        $file_path = $_FILES['blibli_xlsx']['tmp_name'];

        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file_path);
            $sheet = $spreadsheet->getActiveSheet();

            // Loop through the rows in the spreadsheet, starting from the second row
            $highestRow = $sheet->getHighestRow();

            for ($row = 2; $row <= $highestRow; $row++) {
                // Get data from each column
                $product_code = trim($sheet->getCell('A' . $row)->getValue());
                $formatted_name = trim($sheet->getCell('E' . $row)->getValue());
                $sku = trim($sheet->getCell('F' . $row)->getValue());
                $price = trim($sheet->getCell('G' . $row)->getValue());
                $stock = trim($sheet->getCell('I' . $row)->getValue());

                // Use the SKU to check if the product exists
                if (empty($sku)) {
                    error_log("Skipped row $row: SKU is empty");
                    continue;
                }

                $product_id = wc_get_product_id_by_sku($sku);

                if ($product_id) {
                    // Update the existing product
                    $product = wc_get_product($product_id);
                } else {
                    // // Create a new product
                    // $product = new WC_Product();
                    // If the SKU does not exist, skip the product
                    error_log("Skipped row $row: SKU $sku does not exist");
                    continue;
                }

                // Update product data
                // $product->set_status('private'); // Set the product status as private
                // $product->set_name($formatted_name);
                $product->set_stock_quantity($stock);

                // Set the SKU
                $product->set_sku($sku);

                // Update the Blibli price meta field
                update_post_meta($product->get_id(), '_blibli_price', $price);

                // Set the ACF field 'blibli_button' with the formatted URL
                $blibli_button_url = 'https://www.blibli.com/p/product-detail/ps--' . $product_code;
                update_field('button_blibli', $blibli_button_url, $product->get_id());

                // Save the product
                try {
                    $product_id = $product->save();
                    error_log("Product saved: $formatted_name (SKU: $sku, ID: $product_id)");
                    wp_cache_flush();
                } catch (Exception $e) {
                    error_log("Failed to save product: " . $e->getMessage());
                }
            }

            // Redirect to the import page with a success message
            wp_redirect(add_query_arg(['import_status' => 'success', 'import_source' => 'blibli'], admin_url('admin.php?page=blibli-import')));
            exit;
        } catch (Exception $e) {
            error_log("Error during import: " . $e->getMessage());
            wp_redirect(add_query_arg('import_status', 'error', admin_url('admin.php?page=blibli-import')));
            exit;
        }
    } else {
        wp_redirect(add_query_arg('import_status', 'error', admin_url('admin.php?page=blibli-import')));
        exit;
    }
}


function upload_image_from_url_blibli($url)
{
    $upload_dir = wp_upload_dir();
    $image_data = @file_get_contents($url, false, stream_context_create(['http' => ['timeout' => 10]]));

    if ($image_data) {
        $filename = basename($url);
        $file_path = $upload_dir['path'] . '/' . $filename;

        file_put_contents($file_path, $image_data);

        $attachment = array(
            'post_mime_type' => mime_content_type($file_path),
            'post_title'     => sanitize_file_name($filename),
            'post_content'   => '',
            'post_status'    => 'inherit'
        );

        $attach_id = wp_insert_attachment($attachment, $file_path);

        require_once(ABSPATH . 'wp-admin/includes/image.php');
        $attach_data = wp_generate_attachment_metadata($attach_id, $file_path);
        wp_update_attachment_metadata($attach_id, $attach_data);

        return $attach_id;
    } else {
        error_log("Failed to fetch image: $url");
        return false;
    }
}

add_action('admin_notices', 'blibli_import_notices');

function blibli_import_notices()
{
    if (isset($_GET['import_status']) && isset($_GET['import_source']) && $_GET['import_source'] == 'blibli') {
        if ($_GET['import_status'] == 'success') {
            echo '<div class="notice notice-success is-dismissible"><p>Blibli products have been successfully imported.</p></div>';
        } elseif ($_GET['import_status'] == 'error') {
            echo '<div class="notice notice-error is-dismissible"><p>There was an error importing the Blibli products. Please try again.</p></div>';
        }
    }
}
