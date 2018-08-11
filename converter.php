

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
        $this->processDescriptionHeaders();

        if($this->work_mode == 'file') {
            $email = file_get_contents($this->source_filename);
            $this->processEmail($email);
        } else if ($this->work_mode == 'dirs') {
            $this->processDirectory();
        }
        
        fclose($this->output_file_handle);
        
        if ($this->debug == true) {
            $time_end = microtime(true);
            $execution_time = ($time_end - $time_start)/60;
        }

        $this->resultToScreen();

        if ($this->debug == true) {
            echo 'Max memory usage: ' . memory_get_peak_usage(TRUE)/1024/1024 . 'MB' . '<br />';
            echo 'Script Execution time: ' . round(($time_end - $time_start), 4) . ' seconds, ' . round($execution_time, 4) . ' minutes';
        }
        
    }

    public function resultToScreen() {
        if (file_exists($this->output_filename)) {
            echo 'Download file: <a href="http://' . $_SERVER["HTTP_HOST"] . $this->getScriptPath() . '/' .  $this->output_filename.'">http://' . $_SERVER["HTTP_HOST"] . $this->getScriptPath() . $this->output_filename . '</a><br /><br />';
        }
    }

    public function processDirectory( $resource = NULL ) {
        $resource = isset($resource) ? $resource : $this->source_dir;
        
        if ($handle = opendir($resource)) {
            while (false !== ($entry = readdir($handle))) {
                if ($entry != '..' && $entry != '.') {
                    if (is_file($resource . "/" . $entry)) {
                        $email = file_get_contents($resource . "/" . $entry, FILE_USE_INCLUDE_PATH);
                        $this->processEmail($email);
                    } else if (is_dir($resource . "/" . $entry)) {
                        $this->processDirectory($resource . "/" . $entry);
                    }
                }
            }
            closedir($handle);
        }
    }

    public function processDescriptionHeaders(){
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

    public function processEmail($email) {
        $email_contents = $this->separateContent($email);

        if (is_array($email_contents) && count($email_contents) > 1) {
            $email_headers = $email_contents[0];
        } else {
            $email_contents = $this->separateContent($email, "\n\n");            
            if (is_array($email_contents) && count($email_contents) > 1) {
                $email_headers = $email_contents[0];
            } else {
                echo '<br /><br />Exception at record: ' . $email . '<br /><br />';
            }
        }

        if (isset($email_headers) && isset($email_contents)) {
            $processed_headers = $this->processHeaders($email_headers);

            $this->searchForHeaders($processed_headers);

            $this->createCSV($processed_headers);
            
            if (isset($email_contents[1])) {
                $email_body = "";
                if (count($email_contents) > 2) {
                    for ($i=0; $i < count($email_contents); $i++) {
                        if ($i != 0) {
                            $email_body .= $email_contents[$i];
                        }
                    }
                    $this->processContent($email_body);
                } else {
                    $email_body = $email_contents[1];
                    $this->processContent($email_body);
                }
            } else {
                $this->row_content .= $this->field_delimeter . $this->string_delimeter . $this->line_delimeter . $this->string_delimeter;
            }
            $this->writeToFile($this->row_content);
        } else {
            echo '<br /><br />Error at processing record: ' . $email . '<br /><br />';
        }
    }

    public function searchForHeaders($processed_headers) {
        // $processed_headers = implode("\r\n", $processed_headers);
        
        $this->resetFoundValues();

        foreach ($processed_headers as $processed_headers_key => $processed_headers_value) {
            if ($this->searchForHeader('Message-ID', $processed_headers_value)) {
                $this->messageid_found = $this->searchForHeader('Message-ID', $processed_headers_value);
            }
            if ($this->searchForHeader('Date', $processed_headers_value)) {
                $this->date_found = $this->searchForHeader('Date', $processed_headers_value);
            }
            if ($this->searchForHeader('From', $processed_headers_value)) {
                $this->from_found = $this->searchForHeader('From', $processed_headers_value);
            }
            if ($this->searchForHeader('To', $processed_headers_value)) {
                $this->to_found = $this->searchForHeader('To', $processed_headers_value);
            }
            if ($this->searchForHeader('Subject', $processed_headers_value)) {
                $this->subject_found = $this->searchForHeader('Subject', $processed_headers_value);
            }
            if ($this->searchForHeader('Cc', $processed_headers_value)) {
                $this->cc_found = $this->searchForHeader('Cc', $processed_headers_value);
            }
            if ($this->searchForHeader('Mime-Version', $processed_headers_value)) {
                $this->mimeversion_found = $this->searchForHeader('Mime-Version', $processed_headers_value);
            }
            if ($this->searchForHeader('Content-Type', $processed_headers_value)) {
                $this->contentype_found = $this->searchForHeader('Content-Type', $processed_headers_value);
            }
            if ($this->searchForHeader('Content-Transfer-Encoding', $processed_headers_value)) {
                $this->contenttransferencoding_found = $this->searchForHeader('Content-Transfer-Encoding', $processed_headers_value);
            }
            if ($this->searchForHeader('Bcc', $processed_headers_value)) {
                $this->bcc_found = $this->searchForHeader('Bcc', $processed_headers_value);
            }
            if ($this->searchForHeader('X-from', $processed_headers_value)) {
                $this->xfrom_found = $this->searchForHeader('X-from', $processed_headers_value);
            }
            if ($this->searchForHeader('X-to', $processed_headers_value)) {
                $this->xto_found = $this->searchForHeader('X-to', $processed_headers_value);
            }
            if ($this->searchForHeader('X-cc', $processed_headers_value)) {
                $this->xcc_found = $this->searchForHeader('X-cc', $processed_headers_value);
            }
            if ($this->searchForHeader('X-bcc', $processed_headers_value)) {
                $this->xbcc_found = $this->searchForHeader('X-bcc', $processed_headers_value);
            }
            if ($this->searchForHeader('X-folder', $processed_headers_value)) {
                $this->xfolder_found = $this->searchForHeader('X-folder', $processed_headers_value);
            }
            if ($this->searchForHeader('X-origin', $processed_headers_value)) {
                $this->xorigin_found = $this->searchForHeader('X-origin', $processed_headers_value);
            }
            if ($this->searchForHeader('X-fileName', $processed_headers_value)) {
                $this->xfilename_found = $this->searchForHeader('X-fileName', $processed_headers_value);
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

    public function searchForHeader($header, $headers) {
        // return strpos($headers, $header .':');
        if (substr($headers, 0, strlen($header . ":")) == $header . ":") {
            return True;
        } else {
            return False;
        }
    }

    public function processHeaders($email_headers) {
        $cleaned_headers = $this->replaceTabs($email_headers);
        $cleaned_headers = $this->replaceDoubleLines($cleaned_headers);
        $cleaned_headers = $this->replaceSemicolons($cleaned_headers);
        $cleaned_headers = $this->replaceQuotes($cleaned_headers);
        $cleaned_headers = $this->replaceApos($cleaned_headers);
        $cleaned_headers = $this->replaceBackslashes($cleaned_headers);

        $normalized_headers = $this->normalizeHeaders($cleaned_headers);

        $separated_headers = $this->separateHeaders($normalized_headers);
        if (count($separated_headers) <= 1) {
            $separated_headers = $this->separateHeaders($normalized_headers, "\n");
        }

        return $this->trimHeaders($separated_headers);
    }

    public function processHeader($email_header_name, $email_header_value, $field_delimeter = True) {
        $cleaned_header_value = $this->removeHeader($email_header_name, $email_header_value);

        if($cleaned_header_value == "" || $cleaned_header_value == " ") {
            $cleaned_header_value = $this->empty_value;
        }
        
        $cleaned_header_value = trim($cleaned_header_value);
        
        if ($field_delimeter) {
            $this->row_content .= $this->field_delimeter;
        }

        $this->row_content .= $this->string_delimeter . $cleaned_header_value;
    }

    public function processContent($email_content) {
        $email_content = $this->replaceSemicolons($email_content);
        $email_content = $this->replaceQuotes($email_content);
        $email_content = $this->replaceApos($email_content);
        $email_content = $this->replaceBackslashes($email_content);
        $email_content = $this->replaceExtraNewlines($email_content);
        $email_content = $this->replaceNewLines($email_content);
        // $email_content = " ";
        $this->row_content .= $this->field_delimeter . $this->string_delimeter . $email_content . $this->string_delimeter . $this->line_delimeter;
    }

    public function createCSV($processed_headers){
        foreach ($processed_headers as $processed_headers_key => $processed_headers_value) {
            $matches = 0;

            if ( $this->searchForHeader('Message-ID', $processed_headers_value) ) {
                $this->processHeader('Message-ID', $processed_headers_value, False);

                $matches++;

                if ($this->date_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('Date', $processed_headers_value) ) {
                $this->processHeader('Date', $processed_headers_value);
                
                $matches++;

                if ($this->from_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('From', $processed_headers_value) ) {
                $this->processHeader('From', $processed_headers_value);

                $matches++;

                if ($this->to_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('To', $processed_headers_value) ) {
                $this->processHeader('To', $processed_headers_value);
                
                $matches++;
                
                if ($this->subject_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('Subject', $processed_headers_value) ) {
                $this->processHeader('Subject', $processed_headers_value);
                
                $matches++;
                
                if ($this->cc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('Cc', $processed_headers_value) ) {
                $this->processHeader('Cc', $processed_headers_value);
                
                $matches++;
               
                if ($this->mimeversion_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('Mime-Version', $processed_headers_value) ) {
                $this->processHeader('Mime-Version', $processed_headers_value);
                
                $matches++;
                
                if ($this->contentype_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('Content-Type', $processed_headers_value) ) {
                $this->processHeader('Content-Type', $processed_headers_value);
                
                $matches++;
                
                if ($this->contenttransferencoding_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('Content-Transfer-Encoding', $processed_headers_value) ) {
                $this->processHeader('Content-Transfer-Encoding', $processed_headers_value);
                
                $matches++;
                
                if ($this->bcc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('Bcc', $processed_headers_value) ) {
                $this->processHeader('Bcc', $processed_headers_value);
                
                $matches++;
                
                if ($this->xfrom_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('X-from', $processed_headers_value) ) {
                $this->processHeader('X-from', $processed_headers_value);
                
                $matches++;
                
                if ($this->xto_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('X-to', $processed_headers_value) ) {
                $this->processHeader('X-to', $processed_headers_value);
                
                $matches++;
                
                if ($this->xcc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('X-cc', $processed_headers_value) ) {
                $this->processHeader('X-cc', $processed_headers_value);
                
                $matches++;
                
                if ($this->xbcc_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('X-bcc', $processed_headers_value) ) {
                $this->processHeader('X-bcc', $processed_headers_value);
                
                $matches++;
                
                if ($this->xfolder_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('X-folder', $processed_headers_value) ) {
                $this->processHeader('X-folder', $processed_headers_value);
                
                $matches++;
                
                if ($this->xorigin_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('X-origin', $processed_headers_value) ) {
                $this->processHeader('X-origin', $processed_headers_value);
                
                $matches++;
                
                if ($this->xfilename_found === false) {
                    $this->row_content .= $this->string_delimeter . $this->field_delimeter . $this->string_delimeter . $this->empty_value;
                }
            }

            if ( $this->searchForHeader('X-fileName', $processed_headers_value) ) {
                $this->processHeader('X-fileName', $processed_headers_value);
                
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

    public function replaceSemicolons($unified_headers) {
        return str_replace(";" , "&#59;", $unified_headers);
    }

    public function replaceQuotes($unified_headers) {
        return str_replace('"' , "&quot;", $unified_headers);
    }

    public function replaceApos($unified_headers) {
        return str_replace("'" , "&#39;", $unified_headers);
    }

    public function replaceBackslashes($unified_headers) {
        return str_replace("\\" , "&bsol;", $unified_headers);
    }

    public function separateHeaders($email_headers, $separator = "\r\n") {
        return explode($separator, $email_headers);
    }

    public function separateContent($email, $separator = "\r\n\r\n"){
        return explode($separator, $email);
    }
    public function trimHeaders($separated_headers) {
        $trimmed_headers = array();
        foreach($separated_headers as $k => $v) {
            $trimmed_headers[] = trim($v);
        }
        return $trimmed_headers;
    }

    public function replaceTabs($separated_headers) {
        return str_replace("\t", " ", $separated_headers);
    }

    public function normalizeHeaders($cleaned_headers) {
        return str_replace($this->source_headers, $this->target_headers, $cleaned_headers);
    }

    public function removeHeader($header_name, $normalized_header_value){
        return str_replace($header_name .':', "", $normalized_header_value);
    }

    public function replaceDoubleLines($email_headers){
        $email_headers = preg_replace("/\r\n\s+/m", " ", $email_headers);
        return str_replace("  ", " ", $email_headers);
    }

    public function replaceExtraNewlines($email_content){
        return preg_replace("/[\r\n]+/", " ", $email_content);
    }

    public function replaceNewLines($value) {
        $value = str_replace("\r\n", " ", $value);
        return str_replace("\n", "", $value);
    }

    public function writeToFile($content) {
        fwrite($this->output_file_handle, $content);
        $this->row_content = '';
    }

    public function getScriptPath() {
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