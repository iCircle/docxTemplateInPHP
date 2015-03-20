<?php

include_once dirname(__FILE__).'/../../src/DocxTemplate.php';

class ParseXMLTest extends PHPUnit_Framework_TestCase{

    public function testParseXML(){
        $templatePath = dirname(__FILE__).'/template.docx';

        $docxTemplate = new DocxTemplate($templatePath);
        $outputPath = dirname(__FILE__).'/mergedOutput.docx';
        if(file_exists($outputPath)){
            unlink($outputPath);
        }
        $this->assertFalse(file_exists($outputPath));

        //testing merge method
        $docxTemplate->merge(array(),$outputPath);

        $this->assertTrue(file_exists($outputPath));

    }

    public function testDom(){
        $domDocument = new DOMDocument();
        $domDocument->loadXML("<sp><p><c>c1</c><c>c2</c></p></sp>");

        $docElem = $domDocument->documentElement;

        $domElem = $this->parseXML($docElem);

        echo $domDocument->saveXML();


    }

    public function parseXML(DOMElement $domElement){
        if($domElement->tagName == "c"){
            $nextElement = $domElement->nextSibling;
            if(isset($nextElement) && $nextElement){
                foreach($nextElement->childNodes as $nextElementChild){
                    $domElement->appendChild($nextElementChild);
                }
                
            }
        }else{
            foreach($domElement->childNodes as $child){
                $domElement->replaceChild($this->parseXML($child),$child);
            }
        }
        return $domElement;
    }

}




?>
