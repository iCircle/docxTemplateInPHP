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
}




?>
