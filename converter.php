

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
    public $empty_value = 0;
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
    public $messageid_found = False;
    public $date_found = False;
    public $from_found = False;
    public $to_found = False;
    public $subject_found = False;
    public $cc_found = False;
    public $mimeversion_found = False;
    public $contentype_found = False;
    public $contenttransferencoding_found = False;
    public $bcc_found = False;
    public $xfrom_found = False;
    public $xto_found = False;
    public $xcc_found = False;
    public $xbcc_found = False;
    public $xfolder_found = False;
    public $xorigin_found = False;
    public $xfilename_found = False;

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
            $this->process_dir();
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

    public function process_dir( $resource = NULL ) {
        $resource = isset($resource) ? $resource : $this->source_dir;
        
        if ($handle = opendir($resource)) {
            while (false !== ($entry = readdir($handle))) {
                if ($entry != '..' && $entry != '.') {
                    if (is_file($resource . "/" . $entry)) {
                        $email = file_get_contents($resource . "/" . $entry, FILE_USE_INCLUDE_PATH);
                        $this->process_email($email);
                    } else if (is_dir($resource . "/" . $entry)) {
                        $this->process_dir($resource . "/" . $entry);
                    }
                }
            }
            closedir($handle);
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

    public function process_email($email) {
        $email_contents = $this->separate_content($email);

        if (is_array($email_contents) && count($email_contents) > 1) {
            $email_headers = $email_contents[0];
        } else {
            $email_contents = $this->separate_content($email, "\n\n");            
            if (is_array($email_contents) && count($email_contents) > 1) {
                $email_headers = $email_contents[0];
            } else {
                echo '<br /><br />Exception at record: ' . $email . '<br /><br />';
            }
        }

        if (isset($email_headers) && isset($email_contents)) {
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
        } else {
            echo '<br /><br />Error at processing record: ' . $email . '<br /><br />';
        }
    }

    public function search_for_headers($processed_headers) {
        // $processed_headers = implode("\r\n", $processed_headers);
        
        $this->resetFoundValues();

        foreach ($processed_headers as $processed_headers_key => $processed_headers_value) {
            if ($this->search_for_header('Message-ID', $processed_headers_value)) {
                $this->messageid_found = $this->search_for_header('Message-ID', $processed_headers_value);
            }
            if ($this->search_for_header('Date', $processed_headers_value)) {
                $this->date_found = $this->search_for_header('Date', $processed_headers_value);
            }
            if ($this->search_for_header('From', $processed_headers_value)) {
                $this->from_found = $this->search_for_header('From', $processed_headers_value);
            }
            if ($this->search_for_header('To', $processed_headers_value)) {
                $this->to_found = $this->search_for_header('To', $processed_headers_value);
            }
            if ($this->search_for_header('Subject', $processed_headers_value)) {
                $this->subject_found = $this->search_for_header('Subject', $processed_headers_value);
            }
            if ($this->search_for_header('Cc', $processed_headers_value)) {
                $this->cc_found = $this->search_for_header('Cc', $processed_headers_value);
            }
            if ($this->search_for_header('Mime-Version', $processed_headers_value)) {
                $this->mimeversion_found = $this->search_for_header('Mime-Version', $processed_headers_value);
            }
            if ($this->search_for_header('Content-Type', $processed_headers_value)) {
                $this->contentype_found = $this->search_for_header('Content-Type', $processed_headers_value);
            }
            if ($this->search_for_header('Content-Transfer-Encoding', $processed_headers_value)) {
                $this->contenttransferencoding_found = $this->search_for_header('Content-Transfer-Encoding', $processed_headers_value);
            }
            if ($this->search_for_header('Bcc', $processed_headers_value)) {
                $this->bcc_found = $this->search_for_header('Bcc', $processed_headers_value);
            }
            if ($this->search_for_header('X-from', $processed_headers_value)) {
                $this->xfrom_found = $this->search_for_header('X-from', $processed_headers_value);
            }
            if ($this->search_for_header('X-to', $processed_headers_value)) {
                $this->xto_found = $this->search_for_header('X-to', $processed_headers_value);
            }
            if ($this->search_for_header('X-cc', $processed_headers_value)) {
                $this->xcc_found = $this->search_for_header('X-cc', $processed_headers_value);
            }
            if ($this->search_for_header('X-bcc', $processed_headers_value)) {
                $this->xbcc_found = $this->search_for_header('X-bcc', $processed_headers_value);
            }
            if ($this->search_for_header('X-folder', $processed_headers_value)) {
                $this->xfolder_found = $this->search_for_header('X-folder', $processed_headers_value);
            }
            if ($this->search_for_header('X-origin', $processed_headers_value)) {
                $this->xorigin_found = $this->search_for_header('X-origin', $processed_headers_value);
            }
            if ($this->search_for_header('X-fileName', $processed_headers_value)) {
                $this->xfilename_found = $this->search_for_header('X-fileName', $processed_headers_value);
            }
        }
    }

    public function resetFoundValues(){
        $this->messageid_found = False;
        $this->date_found = False;
        $this->from_found = False;
        $this->to_found = False;
        $this->subject_found = False;
        $this->cc_found = False;
        $this->mimeversion_found = False;
        $this->contentype_found = False;
        $this->contenttransferencoding_found = False;
        $this->bcc_found = False;
        $this->xfrom_found = False;
        $this->xto_found = False;
        $this->xcc_found = False;
        $this->xbcc_found = False;
        $this->xfolder_found = False;
        $this->xorigin_found = False;
        $this->xfilename_found = False;
    }

    public function search_for_header($header, $headers) {
        // return strpos($headers, $header .':');
        if (substr($headers, 0, strlen($header . ":")) == $header . ":") {
            return True;
        } else {
            return False;
        }
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
        if (count($separated_headers) <= 1) {
            $separated_headers = $this->separate_headers($normalized_headers, "\n");
        }

        return $this->trim_headers($separated_headers);
    }

    public function process_header($email_header_name, $email_header_value, $field_delimeter = True) {
        $cleaned_header_value = $this->remove_header($email_header_name, $email_header_value);
        if($cleaned_header_value == "" || $cleaned_header_value == " ") {
            $cleaned_header_value = $this->empty_value;
        }
        $cleaned_header_value = trim($cleaned_header_value);
        if ($field_delimeter) {
            $this->row_content .= $this->field_delimeter;
        }
        $this->row_content .= $this->string_delimeter . $cleaned_header_value;
    }

    public function process_content($email_content) {
        $email_content = $this->replace_semicolons($email_content);
        $email_content = $this->replace_quotes($email_content);
        $email_content = $this->replace_apos($email_content);
        $email_content = $this->replace_backslash($email_content);
        $email_content = $this->replace_extra_newlines($email_content);
        $email_content = $this->replace_newlines($email_content);
        // $email_content = " ";
        $this->row_content .= $this->field_delimeter . $this->string_delimeter . $this->replace_newlines($email_content) . $this->string_delimeter . $this->line_delimeter;
    }

    public function create_csv($processed_headers){
        foreach ($processed_headers as $processed_headers_key => $processed_headers_value) {
            $matches = 0;

            if ( $this->search_for_header('Message-ID', $processed_headers_value) ) {
                $this->process_header('Message-ID', $processed_headers_value, False);

                $matches++;

                if ($this->date_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('Date', $processed_headers_value) ) {
                $this->process_header('Date', $processed_headers_value);
                
                $matches++;

                if ($this->from_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('From', $processed_headers_value) ) {
                $this->process_header('From', $processed_headers_value);

                $matches++;

                if ($this->to_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('To', $processed_headers_value) ) {
                $this->process_header('To', $processed_headers_value);
                
                $matches++;
                
                if ($this->subject_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('Subject', $processed_headers_value) ) {
                $this->process_header('Subject', $processed_headers_value);
                
                $matches++;
                
                if ($this->cc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('Cc', $processed_headers_value) ) {
                $this->process_header('Cc', $processed_headers_value);
                
                $matches++;
               
                if ($this->mimeversion_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('Mime-Version', $processed_headers_value) ) {
                $this->process_header('Mime-Version', $processed_headers_value);
                
                $matches++;
                
                if ($this->contentype_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('Content-Type', $processed_headers_value) ) {
                $this->process_header('Content-Type', $processed_headers_value);
                
                $matches++;
                
                if ($this->contenttransferencoding_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('Content-Transfer-Encoding', $processed_headers_value) ) {
                $this->process_header('Content-Transfer-Encoding', $processed_headers_value);
                
                $matches++;
                
                if ($this->bcc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('Bcc', $processed_headers_value) ) {
                $this->process_header('Bcc', $processed_headers_value);
                
                $matches++;
                
                if ($this->xfrom_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('X-from', $processed_headers_value) ) {
                $this->process_header('X-from', $processed_headers_value);
                
                $matches++;
                
                if ($this->xto_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('X-to', $processed_headers_value) ) {
                $this->process_header('X-to', $processed_headers_value);
                
                $matches++;
                
                if ($this->xcc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('X-cc', $processed_headers_value) ) {
                $this->process_header('X-cc', $processed_headers_value);
                
                $matches++;
                
                if ($this->xbcc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('X-bcc', $processed_headers_value) ) {
                $this->process_header('X-bcc', $processed_headers_value);
                
                $matches++;
                
                if ($this->xfolder_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('X-folder', $processed_headers_value) ) {
                $this->process_header('X-folder', $processed_headers_value);
                
                $matches++;
                
                if ($this->xorigin_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('X-origin', $processed_headers_value) ) {
                $this->process_header('X-origin', $processed_headers_value);
                
                $matches++;
                
                if ($this->xfilename_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->search_for_header('X-fileName', $processed_headers_value) ) {
                $this->process_header('X-fileName', $processed_headers_value);
                
                $matches++;
            }

            if ($matches == 0) {
                if(substr($this->row_content, -5) == '";"' . $this->empty_value . '"' ) {
                    $this->row_content = substr($this->row_content, 0, strlen($this->row_content)-5);
                    $this->row_content .= " " . $processed_headers_value . ' ";"' . $this->empty_value . '"';
                } else if (substr($this->row_content, -1) == '"'){
                    $this->row_content = substr($this->row_content, 0, strlen($this->row_content)-1);
                    $this->row_content .= " " . $processed_headers_value . ' "';
                }  
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

    public function separate_headers($email_headers, $separator = "\r\n") {
        return explode($separator, $email_headers);
    }

    public function separate_content($email, $separator = "\r\n\r\n"){
        return explode($separator, $email);
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

    public function replace_newlines($value) {
        $value = str_replace("\r\n", " ", $value);
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