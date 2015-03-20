<?php
/**
 * User: Raghavendra K R
 * Date: 16/3/15
 * Time: 10:05 AM
 */

class DocxTemplate {
    private $template = null;
    private $keyStartChar = '[';
    private $keyEndChar   = ']';

    // for internal Use
    private $incompleteKeyNodes = array();

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
    /**
     * this method normalizes the template keys split over multiple <w:t> and multiple <w:r> elements
     *
     */
    private function normalize(DOMDocument $domDocument){
        $wtElems = $domDocument->getElementsByTagName('w:t');

        foreach($wtElems as $wtElem){
            $textContent = $wtElem->textContent;
            $isIncompleteKeyAtEnd = $this->hasIncompleteKey($textContent);
            while($isIncompleteKeyAtEnd){
                //move the content of next w:t elements to current w:t element to complete the key definition

            }




        }




    }

    private function hasIncompleteKey($text){
        $textChars = str_split($text);
        $isIncompleteKeyAtEnd = false;
        for($i=count($textChars)-1; $i>=0; $i--){
            if($textChars[$i] === $this->$keyStartChar || $textChars[$i] === $this->$keyEndChar){
                // found keyStartChar/keyEndChar, check the \ character behind the keyStartChar/keyEndChar,
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
                    if($textChars[$i] === $this->$keyStartChar){
                        $isIncompleteKeyAtEnd = true;
                    }
                    break;
                }
            }
        }
        return $isIncompleteKeyAtEnd;
    }


    private function parseXMLElement($workingDir,DOMElement $xmlElement,$data){

        $tagName = $xmlElement->tagName;
        switch(strtoupper($tagName)){
            case "W:T":
                //find the template keys and replace it with data
                $keys = $this->getTemplateKeys($xmlElement);

                $xmlElement->nodeValue = "Success1";
                break;
            default:
                if($xmlElement->hasChildNodes()){
                    foreach($xmlElement->childNodes as $childNode){
                        if($childNode->nodeType === XML_ELEMENT_NODE){
                            $newChild = $this->parseXMLElement($workingDir,$childNode,$data);
                            $xmlElement->replaceChild($newChild,$childNode);
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
            return;
        }
        $keys = array();
        $textContent = $wtElement->textContent;
        $incompleteText = '';
        if(count($this->incompleteKeyNodes) > 0){
            // incomplete keys are from different <p> elements , then discard the old incomplete elements
            $firstIncompleteKey = $this->incompleteKeyNodes[0];
            if($firstIncompleteKey->parentNode->parentNode !== $wtElement->parentNode->parentNode){
                $this->log(LOG_WARNING,"incomplete keys in paragraph : Line ".$firstIncompleteKey->parentNode->parentNode->getLineNo());
                $this->incompleteKeyNodes = array();
            }

            foreach($this->incompleteKeyNodes as $incompleteKeyNode){
                //$incompleteKeyNode will be an instance of KeyNode class
                $incompleteText .= $incompleteKeyNode->key();
            }
        }
        $textContent  = $incompleteText.$textContent;

        $textChars = str_split($textContent);



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


}

class KeyNode{
    private $key = null;
    private $keyIndex = 0;
    private $isComplete = false;
    private $wtElement = null;

    function __construct($key,$keyIndex,$isComplete,DOMElement $wtElement){
        $this->key = $key;
        $this->keyIndex = $keyIndex;
        $this->isComplete = $isComplete;
        $this->wtElement = $wtElement;
    }

    /**
     * @return string
     */
    public function key(){
        return $this->key;
    }

    /**
     * @return int
     */
    public function keyIndex(){
        return $this->keyIndex;
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
    public function getWtElement(){
        return $this->wtElement;
    }

}
