<?php

/*
 * hucak 
 * docx to xliff converter..march 2016
 * hasan.ucak@gmail.com
 */

class xliffConverter {

    public $error = '';
    public $error_code = 0;
    //check zip,xmlreader,xmlwriter modules
    //if loaded zip,xmlreader,xmlwriter must be false.
    public $modulCheck = true;
    public $classCheck = true;

    public function __construct() {
        ;
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
     *check necessary  php modules and classes 
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
        }
        return true;
    }

    public function docxToXliff($source, $target, $docxfile) {


        if (!$this->checkFile($docxfile)) {
            return false;
        }


        $xliffData = '<?xml version="1.0" encoding="UTF-8"?>'
                . '<xliff xmlns="urn:oasis:names:tc:xliff:document:1.2" xmlns:its="http://www.w3.org/2005/11/its" xmlns:itsxlf="http://www.w3.org/ns/its-xliff/" xmlns:okp="okapi-framework:xliff-extensions" its:version="2.0" version="1.2">'
                . '<file datatype="x-docx" original="' . basename($docxfile) . '" source-language="' . $source . '" target-language="' . $target . '" tool-id="matecat-converter 1.1.2"><header><reference><internal-file form="base64">';

        $xliffData .=base64_encode(file_get_contents($docxfile)) . '</internal-file></reference></header><body/></file>';

        $xliffData .='<file datatype="x-undefined" original="word/styles.xml" source-language="' . $source . '" target-language="' . $target . '">
<body>
</body>
</file>
<file datatype="x-undefined" original="word/document.xml" source-language="' . $source . '" target-language="' . $target . '">
<body>';

        $zip = new ZipArchive;

        $extractpath = dirname($docxfile) . "/temp";
        if (!file_exists($extractpath))
            mkdir($extractpath, 0755, true);
        if ($zip->open($docxfile) === TRUE) {
            $zip->extractTo($extractpath);
            $zip->close();
        } else {
            LOG::doLog(__METHOD__ . " ZipArchive Error");
            return false;
        }

        $reader = new XMLReader();

        $reader->open($extractpath . "/word/document.xml");

        $idPart = hash('md5', time() . rand(1, 1000));
        $counter = 0;
        while ($reader->read()) {
            switch ($reader->nodeType) {
                case (XMLREADER::ELEMENT):
                    if ($reader->localName == "t") {
                        $xliffData .='<trans-unit id="' . $idPart . '-tu' . ($counter++) . '" xml:space="preserve">
<source xml:lang="' . $source . '">' . $reader->readString() . '</source>
<seg-source><mrk mid="0" mtype="seg">' . $reader->readString() . '</mrk></seg-source>
<target xml:lang="' . $target . '"><mrk mid="0" mtype="seg"></mrk></target>
</trans-unit>';
                    }
            }
        }
        $xliffData .='</body>
</file>
<file datatype="x-undefined" original="word/settings.xml" source-language="' . $source . '" target-language="' . $target . '">
<body>
</body>
</file>
</xliff>';
        file_put_contents($docxfile . ".xlf", $xliffData);
    }

    /*
     * hucak 
     * xliff to docx converter..24 03 2016
     * hasan.ucak@gmail.com
     */

    public function xliffToDocx($data, $orginaldocxfile, $outputdocx) {


        if (!file_exists($orginaldocxfile)) {
            return false;
        }
        $zip = new ZipArchive;

        $extractpath = dirname($orginaldocxfile) . "/temp";
        $extractpath = tempnam("", "docx");

        if (file_exists($extractpath))
            system("rm -rf $extractpath");
        mkdir($extractpath, 0755, true);
        if ($zip->open($orginaldocxfile) === TRUE) {
            $zip->extractTo($extractpath);
            $zip->close();
        } else {
            return false;
        }

        $reader = new XMLReader();
        $writer = new XMLWriter();
        $writer->openURI($extractpath . "/word/tmp_document.xml");
        $writer->startDocument('1.0', 'UTF-8', "yes");

        $reader->XML(file_get_contents($extractpath . "/word/document.xml"));
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

        unlink($extractpath . "/word/document.xml");
        rename($extractpath . "/word/tmp_document.xml", $extractpath . "/word/document.xml");
        $zip = new ZipArchive;
        if ($zip->open($outputdocx, ZIPARCHIVE::CREATE | ZIPARCHIVE::OVERWRITE) === true) {
            $ite = new RecursiveDirectoryIterator($extractpath . "/");
            foreach (new RecursiveIteratorIterator($ite, RecursiveIteratorIterator::LEAVES_ONLY) as $filename => $cur) {
                if (!$cur->isDir()) {
                    $zip->addFile($filename, str_replace($extractpath . "/", "", $filename));
                }
            }
            $zip->close();
        }
        system("rm -rf $extractpath");
    }

}

//Testing
print_r(get_loaded_extensions());

if (!extension_loaded('zip')) {
    print "Yüklü değil";
    exit;
}
?>