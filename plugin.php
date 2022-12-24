<?php
/**
 * Plugin Name: WooCommerce Excel Pricing List
 * Description: Adds a shortcode [download_prices] that when clicked downloads an XLS file with the SKU, name, and price of all the products in the WooCommerce store. Made with the help of Chat GPT.
 * Version: 1.0
 * Author: Unconventional Coding
 * Author URI: https://www.youtube.com/channel/UCqp4QHWfVYd65onoyraiYvg
 */

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

add_shortcode( 'download_prices', 'download_prices_shortcode' );

function download_prices_shortcode( $atts, $content = null ) {
    // Get the link text from the shortcode content, or use a default value
    $text = ! empty( $content ) ? do_shortcode( $content ) : 'Download Prices';

    // Get the URL of the REST API endpoint
    $url = site_url( '/wp-json/my-plugin/v1/products/xls' );

    // Generate the link HTML
    $link = '<a href="' . esc_url( $url ) . '">' . esc_html( $text ) . '</a>';

    return $link;
}

add_action( 'rest_api_init', 'custom_rest_api_endpoint' );

function custom_rest_api_endpoint() {
    register_rest_route( 'my-plugin/v1', '/products/xls', array(
        'methods'  => 'GET',
        'callback' => 'get_products_xls',
    ) );
}

function get_products_xls() {
    require_once __DIR__."/vendor/autoload.php";

    // Set the cache expiration time to 30 minutes
    $cache_expiration = 30 * MINUTE_IN_SECONDS;

    // Get the path and URL of the uploads folder
    $upload_dir = wp_upload_dir();
    $file_path = $upload_dir['basedir'] . '/products.xls';

    // Check if the file exists and has been modified in the past 30 minutes
    $file_modified = file_exists( $file_path ) ? filemtime( $file_path ) : 0;
    $is_cached = ( time() - $file_modified ) < $cache_expiration;

    // If the file is up-to-date, return its URL
    if ( ! $is_cached ) {

        // If the file is outdated or does

        // Get products from the database
        $products = get_posts( array(
            'post_type' => 'product',
            'numberposts' => -1,
        ) );

        // Create a new Spreadsheet object
        $spreadsheet = new Spreadsheet();

        // Set the active sheet
        $spreadsheet->setActiveSheetIndex(0);

        // Add the headers
        $spreadsheet->getActiveSheet()->setCellValue('A1', 'SKU');
        $spreadsheet->getActiveSheet()->setCellValue('B1', 'Name');
        $spreadsheet->getActiveSheet()->setCellValue('C1', 'Price');

        // Set the style for the header cells
        $style = array(
            'font' => array(
                'bold' => true,
            ),
            'fill' => array(
                'fillType' => Fill::FILL_SOLID,
                'color' => array(
                    'rgb' => '808080',
                ),
            ),
            'alignment' => array(
                'horizontal' => Alignment::HORIZONTAL_CENTER,
            ),
        );

        // Apply the style to the header cells
        $spreadsheet->getActiveSheet()->getStyle('A1:C1')->applyFromArray($style);

        // Set the column widths
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20); // Set the width of the "SKU" column to 20 characters
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(40); // Set the width of the "Name" column to 40 characters
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(10); // Set the width of the "Price" column to 10 characters

        // Add the data for each product
        $row = 2;
        foreach ( $products as $product ) {
            $product_object = wc_get_product( $product->ID ); // Create a WP_Product object for the product
            $price = $product_object->get_price(); // Use the WP_Product object to get the price
            if ( ! empty( $price ) ) { // Only include products that have a price
                $spreadsheet->getActiveSheet()->setCellValue( 'A' . $row, $product_object->get_sku() );
                $spreadsheet->getActiveSheet()->setCellValue( 'B' . $row, $product->post_title );
                $spreadsheet->getActiveSheet()->setCellValue( 'C' . $row, $price );
                $row++;
            }
        }

        // Save the XLS file
        $writer = new Xls( $spreadsheet );
        $writer->save( $file_path );
    }

    // Set the proper headers for the XLS file
    header( "Content-Type: application/force-download" );
    header( "Content-Type: application/octet-stream" );
    header( "Content-Type: application/download" );
    header( "Content-Disposition: attachment; filename=products.xls" );
    header( "Content-Transfer-Encoding: binary" );

    readfile( $file_path );

}
