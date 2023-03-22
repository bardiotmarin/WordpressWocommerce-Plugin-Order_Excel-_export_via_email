<?php
/**
 * Plugin Name: WooCommerce Order to Excel by marin
 * Description: Adds an Excel file attachment to WooCommerce order emails
 * Version: 1.0
 * Author: Marin
 * Author URI: monsite.com
 * License: GPL2
 */

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;

if ( defined('CBXPHPSPREADSHEET_PLUGIN_NAME') && file_exists( CBXPHPSPREADSHEET_ROOT_PATH . 'lib/vendor/autoload.php' ) ) {
    //Include PHPExcel
    require_once( CBXPHPSPREADSHEET_ROOT_PATH . 'lib/vendor/autoload.php' );

    //now take instance
    $objPHPExcel = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
}

// Ajouter un filtre pour ajouter un fichier Excel en pièce jointe à l'email de commande
add_filter( 'woocommerce_email_attachments', 'add_excel_attachment_to_order_email', 10, 3 );

function add_excel_attachment_to_order_email( $attachments, $email_id, $order ) {
    if ( 'new_order' === $email_id && $order ) {
        $excel_file = generate_excel_from_order( $order );
        $attachments[] = $excel_file;
    }
    return $attachments;
}

function generate_excel_from_order( $order ) {
    // Répertoire de sauvegarde
    $dir = './backup/';

    // Vérifier si le répertoire existe et est accessible en écriture
    if (!is_dir($dir)) {
        // Répertoire n'existe pas, le créer
        if (!mkdir($dir, 0777, true)) {
            trigger_error('Impossible de créer le répertoire de sauvegarde', E_USER_WARNING);
            return false;
        }
    }
    if (!is_writable($dir)) {
        // Répertoire n'est pas accessible en écriture
        trigger_error('Répertoire de sauvegarde non accessible en écriture', E_USER_WARNING);
        return false;
    }

    $billing_country = $order->get_billing_country();
    $shipping_country = $order->get_shipping_country();
    $currency = $order->get_currency();


    if ( in_array( $billing_country, array( 'FR', 'BE' ) ) || in_array( $shipping_country, array( 'FR', 'BE' ) ) && $currency === 'EUR' ) {
        // Template pour la France et la Belgique (utilisateurs en France ou Belgique et utilisant l'Euro)
        $template_file = 'template-fr.xlsx';
    } elseif ( in_array( $billing_country, array( 'AL', 'AD' ) ) && $currency === 'EUR' ) {
        // Template pour l'Europe (utilisateurs en Europe hors France et Belgique et utilisant l'Euro)
        $template_file = 'template-europe.xlsx';
    } elseif ( $currency === 'USD' ) {
        // Template pour les Etats-Unis (utilisateurs utilisant le dollar américain)
        $template_file = 'template-usa.xlsx';
    } elseif ( $currency === 'CNY' ) {
        // Template pour la Chine (utilisateurs utilisant le RMB chinois)
        $template_file = 'template-zh.xlsx';
    } else {
        // Template par défaut (utilisateurs hors France et Europe utilisant une autre devise que USD ou CNY)
        $template_file = 'template-usa.xlsx';
    }

    // Charger le modèle Excel
    $template_path = plugin_dir_path( __FILE__ ) . $template_file;
    $spreadsheet = IOFactory::load( $template_path );

    // Obtenir la première feuille
    $sheet = $spreadsheet->getActiveSheet();

    // Définir le numéro de commande
    $sheet->setCellValue( 'G1', $order->get_order_number() );

    // Définir la date de la commande


    // Définir la date de la commande
    $sheet->setCellValue( 'A4', $order->get_date_created()->format( 'Y-m-d' ) );

    // Définir les informations client
    //$sheet->setCellValue( 'C4', $order->get_billing_company() );
    $sheet->setCellValue( 'G4', $order->get_billing_first_name() . ' ' . $order->get_billing_last_name() );
    $sheet->setCellValue( 'I4', $order->get_billing_email() );
    //$sheet->setCellValue( 'E5', $order->get_billing_phone() );
    //$sheet->setCellValue( 'C6', $order->get_billing_address_1() . ' ' . $order->get_billing_address_2() );
    //$sheet->setCellValue( 'C7', $order->get_billing_city() );
    //$sheet->setCellValue( 'E7', $order->get_billing_postcode() );
    //$sheet->setCellValue( 'G7', $order->get_billing_country() );

        // Définir les informations de livraison
    //$sheet->setCellValue( 'A9', $order->get_shipping_first_name() . ' ' . $order->get_shipping_last_name() );
    //$sheet->setCellValue( 'C9', $order->get_shipping_address_1() );
    //$sheet->setCellValue( 'C10', $order->get_shipping_address_2() );
    //$sheet->setCellValue( 'E10', $order->get_shipping_city() );
    //$sheet->setCellValue( 'G10', $order->get_shipping_postcode() );
    //$sheet->setCellValue( 'I10', $order->get_shipping_country() );

    // Définir les en-têtes des colonnes de produits
    //$sheet->setCellValue( 'A13', 'SKU' );
    //$sheet->setCellValue( 'B18', 'Nom du produit' );
    //$sheet->setCellValue( 'H18', 'Quantité' );
    //$sheet->setCellValue( 'I18', 'Prix unitaire' );
    //$sheet->setCellValue( 'E13', 'Prix total' );

    // Obtenir la liste des produits de la commande
    $items = $order->get_items();


// Ajouter les produits à la feuille de calcul
$row = 18; // update starting row index for products
foreach ( $items as $item ) {
    // Obtenir les informations du produit
    $product = $item->get_product();
    $sku = $product->get_sku();
    $name = $product->get_name();
    $qty = $item->get_quantity();
    $price = $item->get_total() / $qty;
    $total = $item->get_total();

    // Ajouter les informations du produit à la feuille de calcul
    $sheet->setCellValue( 'G' . $row, $sku );
    $sheet->setCellValue( 'B' . $row, $name ); // update cell reference for product name
    $sheet->setCellValue( 'H' . $row, $qty );
    $sheet->setCellValue( 'I' . $row, $price );
    //$sheet->setCellValue( 'E' . $row, $total );
    
    $row++; // increment row index for the next product
}


// Calculer le total de la commande
$subtotal = $order->get_subtotal();
$shipping_total = $order->get_shipping_total();
$discount_total = $order->get_discount_total();
$tax_total = $order->get_total_tax();
$total = $order->get_total();

// Définir les totaux de la commande
$sheet->setCellValue( 'D' . $row, 'Sous-total:' );
$sheet->setCellValue( 'E' . $row, $subtotal );

$row++;
$sheet->setCellValue( 'D' . $row, 'Livraison:' );
$sheet->setCellValue( 'E' . $row, $shipping_total );

if ( $discount_total > 0 ) {
    $row++;
    $sheet->setCellValue( 'D' . $row, 'Remise:' );
    $sheet->setCellValue( 'E' . $row, -$discount_total );
}

if ( $tax_total > 0 ) {
    $row++;
    $sheet->setCellValue( 'D' . $row, 'Taxe:' );
    $sheet->setCellValue( 'E' . $row, $tax_total );
}

$row++;
$sheet->setCellValue( 'D' . $row, 'Total:' );
$sheet->setCellValue( 'E' . $row, $total );

// Enregistrer le fichier Excel
$filename = $dir . $order->get_order_number() . '.xlsx';
$writer = IOFactory::createWriter( $spreadsheet, 'Xlsx' );
$writer->save( $filename );

return $filename;

?>