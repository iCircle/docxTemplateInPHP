<?php

include_once dirname(__FILE__).'/../../src/DocxTemplate.php';

class ImageReplaceTest extends PHPUnit_Framework_TestCase{

    public function testMerge(){
        $templatePath = dirname(__FILE__).'/template.docx';

        $docxTemplate = new DocxTemplate($templatePath);
        $outputPath = dirname(__FILE__).'/mergedOutput.docx';
        if(file_exists($outputPath)){
            unlink($outputPath);
        }
        $record = array();
        $record["headerImage"] = dirname(__FILE__).'/headerImage.png';
        $record["bodyImage"]   = dirname(__FILE__).'/bodyImage.png';
        $record["footerImage"] = dirname(__FILE__).'/footerImage.png';

        $this->assertFalse(file_exists($outputPath));

        //testing merge method
        $docxTemplate->merge(array("record"=>$record),$outputPath);

        $this->assertTrue(file_exists($outputPath));

    }
}




?>
