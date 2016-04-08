<?php

/*
 * hucak 
 * docx,xlsx,pptx to xliff converter..march 2016
 * hasan.ucak@gmail.com
 */

class xliffConverter {

    public $error = '';
    public $error_code = 0;
//check zip,xmlreader,xmlwriter modules
//if loaded zip,xmlreader,xmlwriter must be false.
    public $modulCheck = true;
    public $classCheck = true;
    private $extractpath = "";

    public function __construct() {
        $this->checkPhpModulClass();
        $this->extractpath = tempnam(sys_get_temp_dir(), "xliffconverter");
    }

    public function extractpath($dirname) {
        $this->extractpath = $dirname;
    }

    private function setError($errStr, $errCode) {
        $this->error = $errStr;
        $this->error_code = $errCode;
    }

    /*
     * Check file exists or not exists
     */

    private function checkFile($filename, $setError = true) {
        if (!file_exists($filename)) {
            if ($setError) {
                $this->setError(" $filename not found", "101");
                return false;
            }
            return false;
        }
        return true;
    }

    /*
     * check necessary  php modules and classes 
     */

    private function checkPhpModulClass() {
        if ($this->modulCheck) {
            if (!extension_loaded('zip')) {
                $this->setError(" Php Zip Module Not Loaded", "102");
                return false;
            }
            if (!extension_loaded('xmlreader')) {
                $this->setError(" Php xmlreader Module Not Loaded", "103");
                return false;
            }
            if (!extension_loaded('xmlwriter')) {
                $this->setError(" Php writer Module Not Loaded", "104");
                return false;
            }
        }
        if ($this->classCheck) {

            if (!class_exists("ZipArchive")) {
                $this->setError(" ZipArchive Class not found", "105");
                return false;
            }
            if (!class_exists("xmlreader")) {
                $this->setError(" xmlreader Class not found", "106");
                return false;
            }
            if (!class_exists("xmlwriter")) {
                $this->setError(" xmlwriter Class not found", "107");
                return false;
            }
            if (!class_exists("tidy")) {
                $this->setError(" Tidy Class not found", "108");
                return false;
            }
        }
        return true;
    }

    /*
     * if has write permisson on directory
     */

    private function checkfPerm($dirname) {
        if (!file_exists($dirname))
            $dirname = dirname($dirname);

        if (is_writable($dirname))
            return true;
        else
            $this->setError(" $dirname not writeable ", "131");
        return false;
    }

    public function officexToXliff($source, $target, $xfile) {

        function xmlToxliff($source, $target, $xmlfile, &$xliffData, $idPart, &$counter) {
            $config = array(
                'indent' => true,
                'input-xml' => true,
                'output-xml' => true,
                'wrap' => false);

            $tidy = new tidy;
            $tidy->parseFile($xmlfile, $config, 'utf8');
            $reader = new XMLReader();
            $reader->XML($tidy->value);

            while ($reader->read()) {
                if ($reader->hasValue && !empty(trim($reader->value))) {
                    $xliffData .='<trans-unit id="' . $idPart . '-tu' . ($counter++) . '" xml:space="preserve">' . PHP_EOL
                            . '<source xml:lang="' . $source . '">' . trim($reader->value) . '</source>' . PHP_EOL
                            . '<seg-source><mrk mid="0" mtype="seg">' . trim($reader->value) . '</mrk></seg-source>' . PHP_EOL
                            . '<target xml:lang="' . $target . '"><mrk mid="0" mtype="seg"></mrk></target>' . PHP_EOL
                            . '</trans-unit>' . PHP_EOL;
                }
            }
            $reader->close();
        }

        libxml_use_internal_errors(true);

        if (!empty($this->error))
            return false;

        if (!$this->checkFile($xfile)) {
            return false;
        }

        $extension = substr($xfile,-4);

        $xliffData = '<?xml version="1.0" encoding="UTF-8"?>'
                . '<xliff xmlns="urn:oasis:names:tc:xliff:document:1.2" xmlns:its="http://www.w3.org/2005/11/its" xmlns:itsxlf="http://www.w3.org/ns/its-xliff/" xmlns:okp="okapi-framework:xliff-extensions" its:version="2.0" version="1.2">'
                . '<file datatype="x-'.$extension.'" original="' . basename($xfile) . '" source-language="' . $source . '" target-language="' . $target . '" tool-id="matecat-converter 1.1.2"><header><reference><internal-file form="base64">';

        $xliffData .=base64_encode(file_get_contents($xfile)) . '</internal-file></reference></header><body/></file>';

        $xliffData .='<file datatype="x-undefined" original="word/styles.xml" source-language="' . $source . '" target-language="' . $target . '">'
                . '<body></body></file><file datatype="x-undefined" original="word/document.xml" source-language="' . $source . '" target-language="' . $target . '"><body>';

        $zip = new ZipArchive;

        if ($this->checkfPerm($this->extractpath))
            return false;

        if (!file_exists($this->extractpath))
            mkdir($this->extractpath, 0755, true);

        if ($zip->open($xfile) === TRUE) {
            $zip->extractTo($this->extractpath);
            $zip->close();
        } else {
            $this->setError(__METHOD__ . " ZipArchive Error", "121");
            return false;
        }

        $xml = array();
        if (file_exists($this->extractpath . "/xl/sharedStrings.xml"))
            $xmlfile[] = $this->extractpath . "/xl/sharedStrings.xml";
        elseif (file_exists($this->extractpath . "/word/document.xml"))
            $xmlfile[] = $this->extractpath . "/word/document.xml";
        elseif (file_exists($this->extractpath . "/ppt/slides/slide1.xml")) {
            $xmlfile = glob($this->extractpath . "/ppt/slides/slide*.xml");
        }

        $idPart = hash('md5', time() . rand(1, 1000));
        $counter = 0;
        foreach ($xmlfile as $filename) {
            xmlToxliff($source, $target, $filename, $xliffData, $idPart, $counter);
        }

        $xliffData .='</body></file><file datatype="x-undefined" original="word/settings.xml" source-language="' . $source . '" target-language="' . $target . '"><body></body></file></xliff>';
        $size = file_put_contents($xfile . ".xlf", $xliffData);
        if ($size == 0) {
            $this->setError("not write to xliff file.", "151");
            return false;
        }
        return true;
    }

    /*
     * hucak 
     * xliff to docx,xlsx,pptx converter..24 03 2016
     * hasan.ucak@gmail.com
     * $data sample array...
     * $data[0]['segment']='.......';
     * $data[0]['translation']='.......';
     * .......
     * .....
     * 
     */

    public function xliffToDocx($data, $orginalxfile, $outputx) {
        libxml_use_internal_errors(true);

        function xlifftoXml($data, $xmlfile, $xmlTmpfile) {
            $config = array(
                'indent' => true,
                'input-xml' => true,
                'output-xml' => true,
                'wrap' => false);

            $tidy = new tidy;
            $tidy->parseFile($xmlfile, $config, 'utf8');

            $reader = new XMLReader();
            $writer = new XMLWriter();
            $writer->openURI($xmlTmpfile);
            $writer->startDocument('1.0', 'UTF-8', "yes");

            $reader->XML($tidy->value);
            while ($reader->read()) {
                $isempty = false;
                switch ($reader->nodeType) {
                    case (XMLREADER::END_ELEMENT):
                        $writer->endElement();
                        break;
                    case (XMLREADER::ELEMENT):
                        $writer->startElement($reader->name);
                        if ($reader->isEmptyElement)
                            $isempty = true;
                        break;
                    case (XMLREADER::TEXT):
                        $key = array_search($reader->value, array_column($data, "segment"));
                        if ($key !== false) {
                            $writer->text($data[$key]['translation']);
                        } else
                            $writer->text($reader->value);
                        break;
                }
                $count = $reader->attributeCount;
                for ($index = 0; $index < $count; $index++) {
                    $reader->moveToAttributeNo($index);
                    $writer->writeAttribute($reader->name, $reader->value);
                }
                if ($isempty) {
                    $writer->endElement();
                }
            }
            $writer->endDocument();
            $reader->close();
        }

        if (!empty($this->error))
            return false;

        if (!$this->checkFile($orginalxfile)) {
            return false;
        }

        if ($this->checkfPerm(dirname($outputx)))
            return false;

        $zip = new ZipArchive;

        if (file_exists($this->extractpath))
            system("rm -rf $this->extractpath");
        mkdir($this->extractpath, 0755, true);
        if ($zip->open($orginalxfile) === TRUE) {
            $zip->extractTo($this->extractpath);
            $zip->close();
        } else {
            return false;
        }

        $xmlfile = array();
        if (file_exists($this->extractpath . "/xl/sharedStrings.xml")) {
            $xmlfile[] = $this->extractpath . "/xl/sharedStrings.xml";
        } elseif (file_exists($this->extractpath . "/word/document.xml")) {
            $xmlfile[] = $this->extractpath . "/word/document.xml";
        } elseif (file_exists($this->extractpath . "/ppt/slides/slide1.xml")) {
            $xmlfile = glob($this->extractpath . "/ppt/slides/slide*.xml");
        }


        foreach ($xmlfile as $xfile) {
            $xmlTmpfile = $xfile . ".tmp" . rand(0, 1000);
            xlifftoXml($data, $xfile, $xmlTmpfile);
            unlink($xfile);
            rename($xmlTmpfile, $xfile);
        }

        $zip = new ZipArchive;
        if ($zip->open($outputx, ZIPARCHIVE::CREATE | ZIPARCHIVE::OVERWRITE) === true) {
            $ite = new RecursiveDirectoryIterator($this->extractpath . "/");
            foreach (new RecursiveIteratorIterator($ite, RecursiveIteratorIterator::LEAVES_ONLY) as $filename => $cur) {
                if (!$cur->isDir()) {
                    $zip->addFile($filename, str_replace($this->extractpath . "/", "", $filename));
                }
            }
            $zip->close();
        } {
            $this->setError(__METHOD__ . " ZipArchive Error", "121");
            return false;
        }

        system("rm -rf $extractpath");
    }

    /*
     * hucak 
     * doc,xls,ppt to docx,xlsx,pptx converter..24 03 2016
     * hasan.ucak@gmail.com
     * require libreoffice 5.x
     * 
     */

    public function officeTOofficex($officeFile) {
        //libreoffice check..
        exec("which libreoffice", $output, $return_var);
        if ($return_var != 0) {
            $this->setError("libreoffice Not install ", "200");
            return false;
        }
        exec("libreoffice --version", $output, $return_var);
        preg_match_all("/^LibreOffice (\d{1})\.(.*)/i", $output[0], $matches, PREG_SET_ORDER);
        if ($matches[0][1] < 5) {
            $this->setError("libreoffice version 5 or greator ", "201");
            return false;
        }


        if (substr($mfile, -4) == ".doc") {
            exec('libreoffice --convert-to docx:"MS Word 2007 XML" ' . $mofficeFile, $output, $return_var);
            if ($return_var != 0) {
                $this->setError("$officeFile file Convert Error:" . var_export($output, true), $return_var);
                return false;
            }
        }

        if (substr($mfile, -4) == ".xls") {
            exec('libreoffice --convert-to xlsx:"MS Excel 2003 XML" ' . $mofficeFile, $output, $return_var);
            if ($return_var != 0) {
                $this->setError("$officeFile file Convert Error:" . var_export($output, true), $return_var);
                return false;
            }
        }

        if (substr($mfile, -4) == ".ppt") {
            exec(' libreoffice --convert-to pptx:"Impress MS PowerPoint 2007 XML" ' . $mofficeFile, $output, $return_var);
            if ($return_var != 0) {
                $this->setError("$officeFile file Convert Error:" . var_export($output, true), $return_var);
                return false;
            }
        }
    }

}

//Testing
?>
