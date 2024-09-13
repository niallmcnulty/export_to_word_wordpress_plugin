<?php
// Include PHPWord library
require_once plugin_dir_path(__FILE__) . '../vendor/autoload.php';

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Shared\Html;

/**
 * Export post content to Word document
 *
 * @param int $post_id The post ID to export.
 */
function etw_export_post_to_word($post_id) {
    $post = get_post($post_id);
    if (!$post) {
        wp_die(__('Post not found', 'export-to-word'));
    }

    // Create a new PHPWord object
    $phpWord = new PhpWord();
    $section = $phpWord->addSection();

    // Add the post title
    $section->addText($post->post_title, array('bold' => true, 'size' => 16));

    // Add the post content
    $html = apply_filters('the_content', $post->post_content);

    // Remove "You may also be interested in these posts"
    $html = preg_replace('/<div class="rp4wp-related-posts">.*?<\/div>/s', '', $html);

    // Allow developers to modify the HTML before export
    $html = apply_filters('etw_export_html', $html, $post);

    Html::addHtml($section, $html, false, false);

    // Add footer
    $footer = $section->addFooter();
    $footer_text = apply_filters('etw_export_footer_text', 'CAPS 123 | caps123.co.za');
    $footer->addText($footer_text, array('size' => 10));

    // Generate the Word document
    $writer = IOFactory::createWriter($phpWord, 'Word2007');

    // Set the appropriate headers for download
    $filename = sanitize_file_name($post->post_title) . '.docx';
    header("Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    header("Content-Disposition: attachment; filename=\"$filename\"");
    header("Cache-Control: max-age=0");

    // Output the file
    $writer->save("php://output");
    exit;
}
