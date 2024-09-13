<?php
/**
 * Plugin Name: Export to Word
 * Plugin URI: https://www.niallmcnulty.com
 * Description: Adds a button to export post content to a Word document.
 * Version: 1.0
 * Author: Niall McNulty
 * Author URI: https://www.niallmcnulty.com
 * Text Domain: export-to-word
 * Domain Path: /languages
 */

// Prevent direct access to this file
if (!defined('ABSPATH')) {
    exit;
}

// Define plugin constants
define('ETW_VERSION', '1.0');
define('ETW_PLUGIN_DIR', plugin_dir_path(__FILE__));
define('ETW_PLUGIN_URL', plugin_dir_url(__FILE__));

// Include PHPWord library
require_once ETW_PLUGIN_DIR . 'vendor/autoload.php';

// Include the main plugin class
require_once ETW_PLUGIN_DIR . 'includes/class-export-to-word.php';

// Initialize the plugin
function etw_init() {
    $plugin = new Export_To_Word();
    $plugin->run();
}
add_action('plugins_loaded', 'etw_init');

// Activation hook
register_activation_hook(__FILE__, 'etw_activate');

function etw_activate() {
    if (version_compare(PHP_VERSION, '7.0', '<')) {
        deactivate_plugins(plugin_basename(__FILE__));
        wp_die(__('Export to Word requires PHP 7.0 or higher.', 'export-to-word'));
    }

    if (!extension_loaded('zip')) {
        deactivate_plugins(plugin_basename(__FILE__));
        wp_die(__('Export to Word requires the ZIP extension to be installed.', 'export-to-word'));
    }
}

// Deactivation hook
register_deactivation_hook(__FILE__, 'etw_deactivate');

function etw_deactivate() {
    // Perform any cleanup if necessary
}

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Shared\Html;

// Add the export button to the post content
function add_export_button($content) {
    if (is_single() && in_the_loop() && is_main_query()) {
        $button = '<form method="post" class="export-to-word-form">
                     <input type="hidden" name="export_to_word" value="1">
                     <input type="submit" value="Export to Word" class="export-to-word-button">
                   </form>';
        return $content . $button;
    }
    return $content;
}
add_filter('the_content', 'add_export_button');

// Handle the export process
function export_to_word() {
    if (isset($_POST['export_to_word']) && is_single()) {
        global $post;
        
        // Create a new PHPWord object
        $phpWord = new PhpWord();
        $section = $phpWord->addSection();
        
        // Add the post title
        $section->addText($post->post_title, array('bold' => true, 'size' => 16));
        
        // Add the post content
        $html = apply_filters('the_content', $post->post_content);
        
        // Remove "You may also be interested in these posts"
        $html = preg_replace('/<div class="rp4wp-related-posts">.*?<\/div>/s', '', $html);
        
        Html::addHtml($section, $html, false, false);
        
        // Add footer
        $footer = $section->addFooter();
        $footer->addText('CAPS 123 | caps123.co.za', array('size' => 10));
        
        // Generate the Word document
        $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
        
        // Set the appropriate headers for download
        header("Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        header("Content-Disposition: attachment; filename=" . sanitize_title($post->post_title) . ".docx");
        header("Cache-Control: max-age=0");
        
        // Output the file
        $objWriter->save("php://output");
        exit;
    }
}
add_action('template_redirect', 'export_to_word');

// Enqueue necessary scripts and styles
function export_to_word_enqueue_scripts() {
    if (is_single()) {
        wp_enqueue_style('export-to-word-style', plugin_dir_url(__FILE__) . 'css/export-to-word.css');
    }
}
add_action('wp_enqueue_scripts', 'export_to_word_enqueue_scripts');
