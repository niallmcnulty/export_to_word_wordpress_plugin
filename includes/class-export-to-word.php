<?php
class Export_To_Word {
    public function run() {
        add_filter('the_content', array($this, 'add_export_button'));
        add_action('template_redirect', array($this, 'handle_export'));
        add_action('wp_enqueue_scripts', array($this, 'enqueue_scripts'));
    }

    public function add_export_button($content) {
        if (is_single() && in_the_loop() && is_main_query()) {
            $nonce = wp_create_nonce('etw_export_nonce');
            $button = sprintf(
                '<form method="post" class="export-to-word-form">
                    <input type="hidden" name="etw_export" value="1">
                    <input type="hidden" name="etw_nonce" value="%s">
                    <button type="submit" class="export-to-word-button">%s</button>
                </form>',
                esc_attr($nonce),
                esc_html__('Export to Word', 'export-to-word')
            );
            return $content . $button;
        }
        return $content;
    }

    public function handle_export() {
        if (isset($_POST['etw_export']) && isset($_POST['etw_nonce'])) {
            if (!wp_verify_nonce($_POST['etw_nonce'], 'etw_export_nonce')) {
                wp_die(__('Security check failed', 'export-to-word'));
            }

            if (!current_user_can('read')) {
                wp_die(__('You do not have permission to export this post', 'export-to-word'));
            }

            $post_id = get_the_ID();
            if (!$post_id) {
                wp_die(__('Invalid post ID', 'export-to-word'));
            }

            try {
                $this->export_post_to_word($post_id);
            } catch (Exception $e) {
                wp_die(sprintf(__('Export failed: %s', 'export-to-word'), $e->getMessage()));
            }
        }
    }

    private function export_post_to_word($post_id) {
        $post = get_post($post_id);
        if (!$post) {
            throw new Exception(__('Post not found', 'export-to-word'));
        }

        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();

        // Add the post title
        $section->addText($post->post_title, array('bold' => true, 'size' => 16));

        // Add the post content
        $html = apply_filters('the_content', $post->post_content);
        $html = preg_replace('/<div class="rp4wp-related-posts">.*?<\/div>/s', '', $html);
        \PhpOffice\PhpWord\Shared\Html::addHtml($section, $html, false, false);

        // Add footer
        $footer = $section->addFooter();
        $footer->addText('CAPS 123 | caps123.co.za', array('size' => 10));

        // Generate the Word document
        $writer = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');

        // Set the appropriate headers for download
        $filename = sanitize_file_name($post->post_title) . '.docx';
        header("Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        header("Content-Disposition: attachment; filename=\"$filename\"");
        header("Cache-Control: max-age=0");

        // Output the file
        $writer->save("php://output");
        exit;
    }

    public function enqueue_scripts() {
        if (is_single()) {
            wp_enqueue_style('export-to-word-style', ETW_PLUGIN_URL . 'css/export-to-word.css', array(), ETW_VERSION);
        }
    }
}
