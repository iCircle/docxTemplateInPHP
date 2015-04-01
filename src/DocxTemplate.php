<?php
/**
 * User: Raghavendra K R
 * Date: 16/3/15
 * Time: 10:05 AM
 */

include_once dirname(__FILE__).'/../vendor/autoload.php';

class DocxTemplate {
    private $template = null;
    private $keyStartChar = '[';
    private $keyEndChar   = ']';
    private $slNoKey = "slNo";
    private $locale = "en_IN";

    // for internal Use
    private $workingDir = null;
    private $workingFile = null;
    private $incompleteKeyNodes = array();

    function __construct($templatePath){
        if(!file_exists($templatePath)){
            throw new Exception("Invalid Template Path");
        }
        $this->template = $templatePath;
    }

    function merge($data, $outputPath, $download = false){
        //open the Archieve to a temp folder

        $this->workingDir = sys_get_temp_dir()."/DocxTemplating";
        if(!file_exists($this->workingDir)){
            mkdir($this->workingDir,0777,true);
        }
        $workingFile = tempnam($this->workingDir,'');
        if($workingFile === FALSE || !copy($this->template,$workingFile)){
            throw new Exception("Error in initializing working copy of the template");
        }
        $this->workingDir = $workingFile."_";
        $zip = new ZipArchive();
        if($zip->open($workingFile) === TRUE){
            $zip->extractTo($this->workingDir);
            $zip->close();
        }else{
            throw new Exception('Failed to extract Template');
        }

        if(!file_exists($this->workingDir)){
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
            if(isset($fileToParse["required"]) && !file_exists($this->workingDir.'/'.$fileToParse["name"])){
                throw new Exception("Can not merge, Template is corrupted");
            }
            if(file_exists($this->workingDir.'/'.$fileToParse["name"])){
                $this->mergeFile($this->workingDir.'/'.$fileToParse["name"],$data);
            }
        }

        // once merge is happened , zip the working directory and rename
        $mergedFile = $this->workingDir.'/output.docx';
        if($zip->open($mergedFile,ZipArchive::CREATE) === FALSE){
            throw new Exception("Error in creating output");
        }

        // Create recursive directory iterator
        $files = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($this->workingDir,FilesystemIterator::SKIP_DOTS),
            RecursiveIteratorIterator::LEAVES_ONLY
        );

        foreach($files as $name=>$file){
                $name = substr($name,strlen($this->workingDir."/"));
                $zip->addFile($file->getRealPath(),$name);
                //echo "\n".$name ."  :  ".$file->getRealPath();
        }
        $zip->close();

        //once merged file is available copy it to $outputPath or write as downloadable file
        if($download != false){
            copy($mergedFile,$outputPath);
        }else{
            $fInfo = new finfo(FILEINFO_MIME);
            $mimeType = $fInfo->file($mergedFile);

            header('Content-Type:'.$mimeType,true);
            header('Content-Length:'.filesize($mergedFile),true);
            header('Content-Disposition: attachment; filename="'.$outputPath.'"',true);
            if(readfile($mergedFile) === FALSE){
                throw new \Exception("Error in reading the file");
            }
        }

        // remove workingDir and workingFile
        unlink($workingFile);
        $this->deleteDir($this->workingDir);
        if($download === true){
            exit;
        }
    }

    private function mergeFile($file,$data){

        $xmlElement = new DOMDocument();
        if($xmlElement->load($file) === FALSE){
            throw new Exception("Error in merging , Template might be corrupted ");
        }

        $this->workingFile = $file;
        $this->parseXMLElement($xmlElement->documentElement,$data);

        if($xmlElement->save($file) === FALSE){
            throw new Exception("Error in creating output");
        }

    }

    private function parseXMLElement(DOMElement $xmlElement,$data){

        $tagName = $xmlElement->tagName;
        switch(strtoupper($tagName)){
            case "W:T":
                //find the template keys and replace it with data
                $keys = $this->getTemplateKeys($xmlElement);
                $textContent = "";
                for($i=0;$i<count($keys);$i++){
                    $key = $keys[$i];
                    if($key->isKey() && $key->isComplete()){
                        $keyOptions = $key->options();
                        $keyName = $key->key();

                        if(array_key_exists("repeat",$keyOptions)){
                            $repeatType = "row";
                            if(array_key_exists("repeatType",$keyOptions)){
                                $repeatType = strtolower($keyOptions["repeatType"]);
                            }
                            switch($repeatType){
                                case "row":
                                    // remove the current key from the w:t textContent
                                    // and add the remaining key's original text unprocessed
                                    for($j=$i+1;$j<count($keys);$j++){
                                        $remainingKey = $keys[$j];
                                        $textContent = $textContent.$remainingKey->originalKey();
                                    }
                                    $this->setTextContent($xmlElement,$textContent);
                                    throw new RepeatRowException($keyName,$keyOptions["repeat"]);
                            }
                        }

                        $keyValue = $this->getValue($keyName,$data);

                        if($keyValue !== false){
                            if(array_key_exists("numberFormat",$keyOptions)){
                                switch(strtolower($keyOptions["numberFormat"])){
                                    case "inwords":
                                        $noToWords = new Numbers_Words();
                                        $keyValue = $noToWords->toCurrency($keyValue,$this->locale);
                                        break;
                                    case "currency":
                                        if($this->locale == "en_IN"){
                                            $keyValue = "".$keyValue;
                                            $keyValue = preg_replace("/\,/","",$keyValue);

                                            $keyValueSplit = preg_split("/\./",$keyValue);
                                            $decimalPart = $keyValueSplit[0];
                                            $fractionPart = "00";
                                            if(count($keyValueSplit) > 1){
                                                $fractionPart = $keyValueSplit[1];
                                            }

                                            $processedDecimalPart = "";
                                            $decimalPart = strrev($decimalPart);
                                            $decimalPart = str_split($decimalPart);
                                            for($k=0;$k<count($decimalPart);$k++){
                                                if($k == 3 || $k == 5 || $k == 7 || $k == 9 || $k == 11 || $k == 13){
                                                    $processedDecimalPart = ",".$processedDecimalPart;
                                                }
                                                $processedDecimalPart = $decimalPart[$k].$processedDecimalPart;
                                            }
                                            if(strlen($fractionPart) == 1){
                                                $fractionPart = $fractionPart."0";
                                            }
                                            $keyValue = $processedDecimalPart.".".$fractionPart;

                                        }else{
                                            $keyValue = number_format($keyValue,2);
                                        }
                                }
                            }
                            $textContent = $textContent.$keyValue;
                        }else{
                            $textContent = $textContent.$this->keyStartChar.$keyName.$this->keyEndChar;
                        }
                    }else{
                        $textContent = $textContent.$key->key();
                    }
                }

                $this->setTextContent($xmlElement,$textContent);
                break;
            case "W:DRAWING":
                $docPrElement = $xmlElement->getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing","docPr")->item(0);
                if($docPrElement !== null){
                    $altText = $docPrElement->getAttribute("descr");
                    if(strlen($altText)>2){
                        $keyNode = new KeyNode($altText,true,true,$docPrElement);
                        $imagePath = $this->getValue($keyNode->key(),$data);

                        $aBlipElem = $xmlElement->getElementsByTagName("blip")->item(0);
                        if(file_exists($imagePath) && $aBlipElem !== null){
                            $resourceId = $aBlipElem->getAttribute("r:embed");
                            $workingFileName = basename($this->workingFile);
                            $relFile = $this->workingDir.'/word/_rels/'.$workingFileName.'.rels';

                            if(file_exists($relFile)){
                                $relDocument = new DOMDocument();
                                $relDocument->load($relFile);
                                $relElements = $relDocument->getElementsByTagName("Relationship");

                                $imageExtn = ".png";

                                $files = array_diff(scandir($this->workingDir.'/word/media'),array(".",".."));
                                $templateImageRelPath = 'media/rImage'.count($files).$imageExtn;
                                $templateImagePath = $this->workingDir.'/word/'.$templateImageRelPath;

                                $newResourceId = "rId".($relElements->length+1);
                                $aBlipElem->setAttribute("r:embed",$newResourceId);

                                $newRelElement = $relDocument->createElement("Relationship");
                                $newRelElement->setAttribute("Id",$newResourceId);
                                $newRelElement->setAttribute("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                                $newRelElement->setAttribute("Target",$templateImageRelPath);

                                $relDocument->documentElement->appendChild($newRelElement);

                                $relDocument->save($relFile);
                                copy($imagePath,$templateImagePath);
                            }
                        }
                    }
                }
                break;
            default:
                if($xmlElement->hasChildNodes()){
                    $childNodes = $xmlElement->childNodes;
                    $childNodesArray = array();
                    foreach($childNodes as $childNode){
                        $childNodesArray[] = $childNode;
                    }
                    foreach($childNodesArray as $childNode){
                        if($childNode->nodeType === XML_ELEMENT_NODE){
                            try{
                                $newChild = $this->parseXMLElement($childNode,$data);
                                $xmlElement->replaceChild($newChild,$childNode);
                            }catch (RepeatTextException $te){
                                //not supported yet
                            }catch (RepeatRowException $re){
                                if(strtoupper($xmlElement->tagName) === "W:TBL"){
                                    $repeatingArray = $this->getValue($re->getKey(),$data);
                                    $nextRow = $childNode->nextSibling;
                                    $repeatingRowElement = $xmlElement->removeChild($childNode);
                                    $repeatingKeyName = $re->getName();
                                    if($repeatingArray && is_array($repeatingArray)){
                                        $slNo = 1;
                                        foreach($repeatingArray as $repeatingData){
                                            $repeatedRowElement = $repeatingRowElement->cloneNode(true);
                                            $repeatingData[$this->slNoKey] = $slNo;
                                            $newData = $data;
                                            $newData[$repeatingKeyName] = $repeatingData;
                                            $generatedRow = $this->parseXMLElement($repeatedRowElement,$newData);
                                            $xmlElement->insertBefore($generatedRow,$nextRow);
                                            $slNo++;
                                        }
                                    }
                                }else{
                                    throw $re;
                                }
                            }
                        }
                    }
                }
        }

        return $xmlElement;
    }

    /**
     * @param DOMElement $wtElement <w:t> element in the document xml,
     * this method should be called sequentially for all the <w:t> elements in the order they appear in the document xml
     *
     */
    private function getTemplateKeys(DOMElement $wtElement){
        if(strtoupper($wtElement->tagName) != "W:T"){
            $this->log(LOG_ALERT,"Invalid element for finding template keys : Line ".$wtElement->getLineNo());
            return false;
        }
        $keys = array();
        $textContent = $wtElement->textContent;
        $incompleteText = '';
        if(count($this->incompleteKeyNodes) > 0){
            // incomplete keys are from different <p> elements , then discard the old incomplete elements
            $firstIncompleteKey = $this->incompleteKeyNodes[0];
            if($firstIncompleteKey->element()->parentNode->parentNode !== $wtElement->parentNode->parentNode){
                $this->log(LOG_WARNING,"incomplete keys in paragraph : Line ".$firstIncompleteKey->element()->parentNode->parentNode->getLineNo());
                $this->incompleteKeyNodes = array();
            }

            foreach($this->incompleteKeyNodes as $incompleteKeyNode){
                //$incompleteKeyNode will be an instance of KeyNode class
                $incompleteText .= $incompleteKeyNode->key();
            }
        }
        $textContent  = $incompleteText.$textContent;

        $textChars = str_split($textContent);
        $key = null;
        $nonKey = "";
        for($i=0;$i<count($textChars);$i++){
            if($textChars[$i] === $this->keyStartChar || $textChars[$i] === $this->keyEndChar){
                // found keyStartChar/keyEndChar check the \ character behind the keyStartChar/keyEndChar
                $j = $i-1;
                for(; $j>= 0;$j--){
                    if($textChars[$j] != "\\"){
                        break;
                    }
                }
                if(($i-$j)%2){
                    // if i-j is odd ,
                    // then there are even numbers of \ chars behind found keyStartChar/keyEndChar
                    // so keyStartChar/keyEndChar is not escaped and hence valid
                    if($textChars[$i] === $this->keyStartChar){
                        //found keyStartChar
                        if($nonKey !== ""){
                            $keyNode = new KeyNode($nonKey,false,true,$wtElement);
                            $keys[] = $keyNode;
                        }
                        if($key != null){
                            $keyNode = new KeyNode($key,false,true,$wtElement);
                            $keys[] = $keyNode;
                        }
                        $key = $textChars[$i];
                        $nonKey = "";
                    }else{
                        //found keyEndChar
                        if($key !== null){
                            $key = $key.$textChars[$i];
                            $keyNode = new KeyNode($key,true,true,$wtElement);
                            $keys[] = $keyNode;
                            $key = null;
                            $nonKey = "";
                        }else{
                            $nonKey = $nonKey.$textChars[$i];
                        }
                    }
                    continue;
                }

            }
            //neither keyStartChar nor keyEndChar
            if($key !== null){
                // if a key is started, append to it
                $key = $key.$textChars[$i];
            }else{
                $nonKey = $nonKey.$textChars[$i];
            }
        }

        if($key !== null){
            $openKey = new KeyNode($key,true,false,$wtElement);
        }
        if($nonKey !== ""){
            $openText = new KeyNode($nonKey,false,true,$wtElement);
        }

        $incompleteKeys = false;
        if(count($this->incompleteKeyNodes) > 0){
            $incompleteKeys = true;
        }
        if($incompleteKeys && (!isset($openKey) || (isset($openKey) && count($keys) > 0))){
            // if there were incomplete keys and found one or more complete keys in current textContent
            // copy the incomplete keys content to current w:t element
            for($i = count($this->incompleteKeyNodes)-1;$i>=0;$i--){
                $incompleteKeyNode = $this->incompleteKeyNodes[$i];
                $incompleteKeyElement = $incompleteKeyNode->element();
                $incompleteKey = $incompleteKeyNode->key();

                //delete content from the incompleteKeyElement
                $incompleteKeyElementContent = $incompleteKeyElement->textContent;
                $incompleteKeyElementContent = substr($incompleteKeyElementContent,0,strlen($incompleteKeyElementContent)-strlen($incompleteKey));
                if($this->endsWith($incompleteKeyElementContent," ")){
                    $incompleteKeyElement->setAttribute("xml:space","preserve");
                }
                $this->setTextContent($incompleteKeyElement,$incompleteKeyElementContent);

                //add incomplete key to this wtElement
                $thisTextContent = $wtElement->textContent;
                $this->setTextContent($wtElement,$incompleteKey.$thisTextContent);
            }
            $this->incompleteKeyNodes = array();
        }

        if(isset($openKey) && (!$incompleteKeys || ($incompleteKeys && count($keys) > 0))){
            $this->incompleteKeyNodes[] = $openKey;
            $keys[] = $openKey;
        }

        if(isset($openKey) && $incompleteKeys && count($keys) == 0){
            $thisTextAsKeyNode = new KeyNode($wtElement->textContent,true,false,$wtElement);
            $this->incompleteKeyNodes[] = $thisTextAsKeyNode;
            $keys[] = $thisTextAsKeyNode;
        }

        if(isset($openText)){
            $keys[]= $openText;
        }
        return $keys;
    }

    private function log($level,$message){
        echo $message;
    }

    private function startsWith($haystack, $needle) {
        // search backwards starting from haystack length characters from the end
        return $needle === "" || strrpos($haystack, $needle, -strlen($haystack)) !== FALSE;
    }
    private function endsWith($haystack, $needle) {
        // search forward starting from end minus needle length characters
        return $needle === "" || (($temp = strlen($haystack) - strlen($needle)) >= 0 && strpos($haystack, $needle, $temp) !== FALSE);
    }

    private function getValue($key,$data){
        $keyParts = preg_split('/\./',$key);
        $keyValue = $data;
        foreach($keyParts as $keyPart){
            $keyPart = trim($keyPart);
            if(is_array($keyValue) && array_key_exists($keyPart,$keyValue)){
                $keyValue = $keyValue[$keyPart];
            }else{
                $keyValue = false;
                break;
            }
        }
        return $keyValue;
    }

    private function setTextContent(DOMNode $node,$value){
        $node->nodeValue = "";
        return $node->appendChild($node->ownerDocument->createTextNode($value));
    }

    static public function deleteDir($dirPath){
        if(is_dir($dirPath)){
            $files = array_diff(scandir($dirPath), array('..', '.'));
            foreach($files as $file){
                self::deleteDir($dirPath.'/'.$file);
            }
            rmdir($dirPath);
        }else{
            unlink($dirPath);
        }
    }
}

class KeyNode{
    private $key = null;
    private $originalKey = null;
    private $isKey = false;
    private $isComplete = false;
    private $element = null;
    private $options = array();

    function __construct($key,$isKey,$isComplete,DOMElement $element){
        $this->key = $key;
        $this->originalKey = $key;
        $this->isKey = $isKey;
        $this->isComplete = $isComplete;
        $this->element = $element;

        if($this->isKey && $this->isComplete){
            //parse the complete key to extract options
            $options = preg_split("/;/",substr($this->key,1,strlen($this->key)-2));
            if(count($options) > 1){
                for($i=1;$i<count($options);$i++){
                    $option = preg_split("/=/",$options[$i]);
                    $optionName = trim($option[0]);
                    $optionValue = true;
                    if(count($option) > 1){
                        $optionValue = trim($option[1]);
                    }
                    $this->options[$optionName] = $optionValue;
                }
            }
            $this->key = trim($options[0]);
            //echo "\n".$this->key;
        }

    }

    /**
     * @return string
     */
    public function key(){
        return $this->key;
    }

    /**
     * @return string
     */
    public function originalKey(){
        return $this->originalKey;
    }

    /**
     * @return int
     */
    public function isKey(){
        return $this->isKey;
    }

    /**
     * @return boolean
     */
    public function isComplete(){
        return $this->isComplete;
    }

    /**
     * @return DOMElement
     */
    public function element(){
        return $this->element;
    }

    /**
     * @return array
     */
    public function options(){
        return $this->options;
    }

}

// Exceptions to handle repetition

class RepeatTextException extends Exception{
    private $name = null;
    private $key = null;

    function __construct($name,$key){
        $this->name = $name;
        $this->key = $key;
    }

    public function getKey(){
        return $this->key;
    }

    public function getName(){
        return $this->name;
    }
}

class RepeatParagraphException extends Exception{
    private $name = null;
    private $key = null;

    function __construct($name,$key){
        $this->name = $name;
        $this->key = $key;
    }

    public function getKey(){
        return $this->key;
    }

    public function getName(){
        return $this->name;
    }
}

class RepeatRowException extends Exception{
    private $name = null;
    private $key = null;

    function __construct($name,$key){
        $this->name = $name;
        $this->key = $key;
    }

    public function getKey(){
        return $this->key;
    }

    public function getName(){
        return $this->name;
    }
}
