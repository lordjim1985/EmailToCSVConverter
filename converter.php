

<?php

class EmailToCSVConverter {
    // Sets standard parameters for converter
    public $field_delimeter = ';';
    public $string_delimeter = '"';
    public $output_filename = 'result.csv';
    public $source_filename = 'email.txt';
    public $line_delimeter = "\r\n";
    public $output_file_handle;
    public $file_headers = "";
    public $file_contents = "";
    public $source_file_handle;
    public $row_content;
    public $row_content_text;
    public $work_mode = "dirs";
    public $source_dir = "maildir";
    public $debug = true;
    public $source_headers = array('Message-ID: ', 'Date: ', 'From: ', 'To: ', 'Subject: ', 'Cc: ', 'Mime-Version: ', 
            'Content-Type: ', 'Content-Transfer-Encoding: ', 'Bcc: ', 'X-From: ', 'X-To: ', 'X-cc: ', 'X-bcc: ', 'X-Folder: ', 
            'X-Origin: ', 'X-FileName: ');
    public $target_headers = array('Message-ID: ', 'Date: ', 'From: ', 'To: ', 'Subject: ', 'Cc: ', 'Mime-Version: ', 
            'Content-Type: ', 'Content-Transfer-Encoding: ', 'Bcc: ', 'X-from: ', 'X-to: ', 'X-cc: ', 'X-bcc: ', 'X-folder: ', 
            'X-origin: ', 'X-fileName: ');
    public $desc_headers = array('Message-ID', 'Date', 'From', 'To', 'Subject', 'Cc', 'Mime-Version', 
            'Content-Type', 'Content-Transfer-Encoding', 'Bcc', 'X-from', 'X-to', 'X-cc', 'X-bcc', 'X-folder', 
            'X-origin', 'X-fileName', 'Message');
    public $messageid_found;
    public $date_found;
    public $from_found;
    public $to_found;
    public $subject_found;
    public $cc_found;
    public $mimeversion_found;
    public $contentype_found;
    public $contenttransferencoding_found;
    public $bcc_found;
    public $xfrom_found;
    public $xto_found;
    public $xcc_found;
    public $xbcc_found;
    public $xfolder_found;
    public $xorigin_found;
    public $xfilename_found;

    public function main() {
        // Main object
        if ($this->debug == true) {
            $time_start = microtime(true);
        }

        if (isset($_POST['emailPath'])) {

            if (is_file($_POST['emailPath'])) {
                $this->work_mode = 'file';
                $this->source_filename = $_POST['emailPath'];
            }

            if (is_dir($_POST['emailPath'])) {
                $this->work_mode = 'dirs';
                $this->source_dir = $_POST['emailPath'];
            }
        }

        $this->output_file_handle = fopen($this->output_filename, "w+");
        $this->process_description_headers();

        if($this->work_mode == 'file') {
            $email = file_get_contents($this->source_filename);
            $this->process_email($email);
        } else if ($this->work_mode == 'dirs') {
            $this->read_dir();
        }
        
        fclose($this->output_file_handle);
        
        if ($this->debug == true) {
            $time_end = microtime(true);
            $execution_time = ($time_end - $time_start)/60;
        }

        $this->result_to_screen();
        if ($this->debug == true) {
            echo 'Max memory usage: ' . memory_get_peak_usage(TRUE)/1024/1024 . 'MB' . '<br />';
            echo 'Script Execution time: ' . round(($time_end - $time_start), 4) . ' seconds, ' . round($execution_time, 4) . ' minutes';
        }
        
    }

    public function result_to_screen() {
        if (file_exists($this->output_filename)) {
            echo 'Download file: <a href="http://' . $_SERVER["HTTP_HOST"] . $this->get_script_path() . '/' .  $this->output_filename.'">http://' . $_SERVER["HTTP_HOST"] . $this->get_script_path() . $this->output_filename . '</a><br /><br />';
        }
    }

    public function process_description_headers(){
        // Set description headers for CSV
        for($i=0;$i < count($this->desc_headers);$i++) {
            if($i != ( count($this->desc_headers) - 1 ) ) {
                $this->file_headers .= $this->string_delimeter . $this->desc_headers[$i] . $this->string_delimeter . $this->field_delimeter;
            } else {
                $this->file_headers .= $this->string_delimeter . $this->desc_headers[$i] . $this->string_delimeter . $this->line_delimeter;
            }

        }

        $this->row_content .= $this->file_headers;
    }

    public function read_dir( $resource = NULL ) {
        $resource = isset($resource) ? $resource : $this->source_dir;
        
        if ($handle = opendir($resource)) {
            while (false !== ($entry = readdir($handle))) {
                if ($entry != '..' && $entry != '.') {
                    if (is_file($resource . "/" . $entry)) {
                        $email = file_get_contents($resource . "/" . $entry, FILE_USE_INCLUDE_PATH);
                        $this->process_email($email);
                    } else if (is_dir($resource . "/" . $entry)) {
                        $this->read_dir($resource . "/" . $entry);
                    }
                }
            }
            closedir($handle);
        }
    }

    public function process_email($email) {
        $email_contents = $this->separate_content($email);
        $email_headers = $email_contents[0];

        $processed_headers = $this->process_headers($email_headers);
        
        $this->search_for_headers($processed_headers);
        $this->create_csv($processed_headers);
        
        if (isset($email_contents[1])) {
            $email_body = "";
            if (count($email_contents) > 2) {
                for ($i=0; $i < count($email_contents); $i++) {
                    if ($i != 0) {
                        $email_body .= $email_contents[$i];
                    }
                }
                $this->process_content($email_body);
            } else {
                $email_body = $email_contents[1];
                $this->process_content($email_body);
            }
        } else {
            $this->row_content .= $this->field_delimeter . $this->string_delimeter . $this->line_delimeter . $this->string_delimeter;
        }
        $this->write_to_file($this->row_content);
    }

    public function search_for_headers($processed_headers) {
        $processed_headers = implode("\r\n", $processed_headers);

        $this->messageid_found = $this->search_for_header('Message-ID', $processed_headers);
        $this->date_found = $this->search_for_header('Date', $processed_headers);
        $this->from_found = $this->search_for_header('From', $processed_headers);
        $this->to_found = $this->search_for_header('To', $processed_headers);
        $this->subject_found = $this->search_for_header('Subject', $processed_headers);
        $this->cc_found = $this->search_for_header('Cc', $processed_headers);
        $this->mimeversion_found = $this->search_for_header('Mime-Version', $processed_headers);
        $this->contentype_found = $this->search_for_header('Content-Type', $processed_headers);
        $this->contenttransferencoding_found = $this->search_for_header('Content-Transfer-Encoding', $processed_headers);
        $this->bcc_found = $this->search_for_header('Bcc', $processed_headers);
        $this->xfrom_found = $this->search_for_header('X-from', $processed_headers);
        $this->xto_found = $this->search_for_header('X-to', $processed_headers);
        $this->xcc_found = $this->search_for_header('X-cc', $processed_headers);
        $this->xbcc_found = $this->search_for_header('X-bcc', $processed_headers);
        $this->xfolder_found = $this->search_for_header('X-folder', $processed_headers);
        $this->xorigin_found = $this->search_for_header('X-origin', $processed_headers);
        $this->xfilename_found = $this->search_for_header('X-fileName', $processed_headers);
    }

    public function search_for_header($header, $headers) {
        return strpos($headers, $header .':');
    }

    public function process_headers($email_headers) {
        $cleaned_headers = $this->replace_tabs($email_headers);
        $cleaned_headers = $this->replace_doublelines($cleaned_headers);
        
        $cleaned_headers = $this->replace_semicolons($cleaned_headers);
        $cleaned_headers = $this->replace_quotes($cleaned_headers);
        $cleaned_headers = $this->replace_apos($cleaned_headers);
        $cleaned_headers = $this->replace_backslash($cleaned_headers);

        $normalized_headers = $this->normalize_headers($cleaned_headers);

        $separated_headers = $this->separate_headers($normalized_headers);      
        return $this->trim_headers($separated_headers);
    }

    public function process_content($email_content) {
        $email_content = $this->replace_semicolons($email_content);
        $email_content = $this->replace_quotes($email_content);
        $email_content = $this->replace_apos($email_content);
        $email_content = $this->replace_backslash($email_content);
        $email_content = $this->replace_extra_newlines($email_content);
        // $email_content = " ";
        $this->row_content .= $this->field_delimeter . $this->string_delimeter . $this->remove_newlines($email_content) . $this->string_delimeter . $this->line_delimeter;
    }

    public function create_csv($processed_headers){
        foreach ($processed_headers as $processed_headers_key => $processed_headers_value) {
            $matches = 0;

            if ( strpos($processed_headers_value, 'Message-ID:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('Message-ID', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->date_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'Date:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('Date', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->from_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'From:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('From', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->to_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'To:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('To', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->subject_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'Subject:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('Subject', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->cc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'Cc:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('Cc', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->mimeversion_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'Mime-Version:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('Mime-Version', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->contentype_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'Content-Type:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('Content-Type', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->contenttransferencoding_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'Content-Transfer-Encoding:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('Content-Transfer-Encoding', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->bcc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'Bcc:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('Bcc', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }                
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->xfrom_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'X-from:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('X-from', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->xto_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'X-to:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('X-to', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->xcc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'X-cc:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('X-cc', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->xbcc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'X-bcc:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('X-bcc', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->xfolder_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'X-folder:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('X-folder', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->xorigin_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'X-origin:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('X-origin', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
                if ($this->xfilename_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . 0;
                }
            }

            if ( strpos($processed_headers_value, 'X-fileName:', 0) !== false ){
                $cleaned_header_value = $this->remove_header('X-fileName', $processed_headers_value);
                if($cleaned_header_value == "" || $cleaned_header_value == " ") {
                    $cleaned_header_value = 0;
                }
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . trim($cleaned_header_value);
                $matches++;
            }

            if ($matches == 0) {
                $this->row_content = substr($this->row_content, 0, strlen($this->row_content)-2);
                $this->row_content .= $this->remove_newlines($processed_headers_value);
            } else {
                $this->row_content .= $this->string_delimeter;
            }

        }
    }

    public function replace_semicolons($unified_headers) {
        return str_replace(";" , "&#59;", $unified_headers);
    }

    public function replace_quotes($unified_headers) {
        return str_replace('"' , "&quot;", $unified_headers);
    }

    public function replace_apos($unified_headers) {
        return str_replace("'" , "&#39;", $unified_headers);
    }

    public function replace_backslash($unified_headers) {
        return str_replace("\\" , "&bsol;", $unified_headers);
    }

    public function separate_headers($email_headers) {
        return explode("\r\n", $email_headers);
    }

    public function separate_content($email){
        return explode("\r\n\r\n", $email);
    }

    public function trim_headers($separated_headers) {
        $trimmed_headers = array();
        foreach($separated_headers as $k => $v) {
            $trimmed_headers[] = trim($v);
        }
        return $trimmed_headers;
    }

    public function replace_tabs($separated_headers) {
        return str_replace("\t", " ", $separated_headers);
    }

    public function normalize_headers($cleaned_headers) {
        return str_replace($this->source_headers, $this->target_headers, $cleaned_headers);
    }

    public function remove_header($header_name, $normalized_header_value){
        return str_replace($header_name .':', "", $normalized_header_value);
    }

    public function replace_doublelines($email_headers){
        $email_headers = preg_replace("/\r\n\s+/m", " ", $email_headers);
        return str_replace("  ", " ", $email_headers);
    }

    public function replace_extra_newlines($email_content){
        return preg_replace("/[\r\n]+/", " ", $email_content);
    }

    public function remove_newlines($value) {
        $value = str_replace("\r\n", "", $value);
        return str_replace("\n", "", $value);
    }

    public function write_to_file($content) {
        fwrite($this->output_file_handle, $content);
        $this->row_content = '';
    }

    public function get_script_path() {
        $script_path_elements = explode("/", $_SERVER['REQUEST_URI']);
        $script_path = '';
        for ($i=0;$i<count($script_path_elements);$i++) {
            if ($i != ( count($script_path_elements) - 1 ) ) {
                $script_path .= $script_path_elements[$i] . '/';
            }
        }
        return $script_path;
    }
}

if (isset($_POST['emailPath']) && $_POST['emailPath'] != '') {
    if (file_exists($_POST['emailPath'])){
        $converter = new EmailToCSVConverter();
        $converter->main();
    } else {
        echo 'Requested file does not exist.';
    }
} else {
    echo 'No filename or directory name given.';
}

?>