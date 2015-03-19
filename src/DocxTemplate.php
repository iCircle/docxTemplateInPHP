<?php
/**
 * User: Raghavendra K R
 * Date: 16/3/15
 * Time: 10:05 AM
 */

class DocxTemplate {
    private $template = null;

    function __construct($templatePath){
        if(!file_exists($templatePath)){
            throw new Exception("Invalid Template Path");
        }
        $this->template = $templatePath;
    }

    function merge($data, $outputPath = null){
        //open the Archieve to a temp folder

        $workingDir = sys_get_temp_dir()."/DocxTemplating";
        if(!file_exists($workingDir)){
            mkdir($workingDir,0777,true);
        }
        $workingFile = tempnam($workingDir,'');
        if($workingFile === FALSE || !copy($this->template,$workingFile)){
            throw new Exception("Error in initializing working copy of the template");
        }
        $workingDir = $workingFile."_";
        $zip = new ZipArchive();
        if($zip->open($workingFile) === TRUE){
            $zip->extractTo($workingDir);
            $zip->close();
        }else{
            throw new Exception('Failed to extract Template');
        }

        if(!file_exists($workingDir)){
            throw new Exception('Failed to extract Template');
        }

        $filesToParse = array(
            array("name"=>"word/document.xml","required"=>true),
            array("name"=>"word/header1.xml"),
            array("name"=>"word/header2.xml"),
            array("name"=>"word/header3.xml"),
            array("name"=>"word/footer1.xml"),
            array("name"=>"word/footer2.xml"),
            array("name"=>"word/footer3.xml"),
            array("name"=>"word/footnotes.xml"),
            array("name"=>"word/endnotes.xml")
        );

        foreach($filesToParse as $fileToParse){
            if(isset($fileToParse["required"]) && !file_exists($workingDir.'/'.$fileToParse["name"])){
                throw new Exception("Can not merge, Template is corrupted");
            }
            if(file_exists($workingDir.'/'.$fileToParse["name"])){
                $this->mergeFile($workingDir,$workingDir.'/'.$fileToParse["name"],$data);
            }
        }

        // once merge is happened , zip the working directory and rename
        $mergedFile = $workingDir.'/output.docx';
        if($zip->open($mergedFile,ZipArchive::CREATE) === FALSE){
            throw new Exception("Error in creating output");
        }

        // Create recursive directory iterator
        $files = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($workingDir),
            RecursiveIteratorIterator::LEAVES_ONLY
        );

        foreach($files as $name=>$file){
            if(!$this->endsWith($name,".") && !$this->endsWith($name,"..")){
                $name = substr($name,strlen($workingDir."/"));
                $zip->addFile($file->getRealPath(),$name);
                //echo "\n".$name ."  :  ".$file->getRealPath();
            }
        }
        $zip->close();

        //once merged file is available copy it to $outputPath or write as downloadable file
        if(isset($outputPath)){
            copy($mergedFile,$outputPath);
        }else{

        }


        // remove workingDir and workingFile
        unlink($workingFile);
        //rmdir($workingDir);

    }

    private function mergeFile($workingDir,$file,$data){

        $xmlElement = new DOMDocument();
        if($xmlElement->load($file) === FALSE){
            throw new Exception("Error in merging , Template might be corrupted ");
        }

        $this->parseXMLElement($workingDir,$xmlElement->documentElement,$data);

        if($xmlElement->save($file) === FALSE){
            throw new Exception("Error in creating output");
        }

    }

    private function parseXMLElement($workingDir,DOMElement $xmlElement,$data){

        $returnObj = array();

        $tagName = $xmlElement->tagName;
        switch(strtoupper($tagName)){
            case "W:T":
                $xmlElement->nodeValue = "Success1";
                $returnObj["status"]="replace";
                $returnObj["element"]=$xmlElement;
                break;
            default:
                if($xmlElement->hasChildNodes()){
                    $status = "none";
                    foreach($xmlElement->childNodes as $childNode){
                        if($childNode->nodeType === XML_ELEMENT_NODE){
                            $parsedElement = $this->parseXMLElement($workingDir,$childNode,$data);
                            switch($parsedElement["status"]){
                                case "replace":
                                    $xmlElement->replaceChild($parsedElement["element"],$childNode);
                                    $status = "replace";
                                    break;
                                case "mergeNext":

                            }
                        }
                    }
                    $returnObj["status"] = $status;
                    $returnObj["element"] = $xmlElement;
                }else{
                    $returnObj["status"] = "none";
                    $returnObj["element"] = $xmlElement;
                }
        }

        return $returnObj;
    }

    private function startsWith($haystack, $needle) {
        // search backwards starting from haystack length characters from the end
        return $needle === "" || strrpos($haystack, $needle, -strlen($haystack)) !== FALSE;
    }
    private function endsWith($haystack, $needle) {
        // search forward starting from end minus needle length characters
        return $needle === "" || (($temp = strlen($haystack) - strlen($needle)) >= 0 && strpos($haystack, $needle, $temp) !== FALSE);
    }


}
