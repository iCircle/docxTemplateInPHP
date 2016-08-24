<?php

use icircle\Template\Docx\DocxTemplate;

include_once dirname(__FILE__).'/../../src/DocxTemplate.php';

class CompleteTest extends PHPUnit_Framework_TestCase{

    public function testMerge(){
        $templatePath = dirname(__FILE__).'/template.docx';

        $docxTemplate = new DocxTemplate($templatePath);
        $outputPath = dirname(__FILE__).'/mergedOutput.docx';
        if(file_exists($outputPath)){
            unlink($outputPath);
        }
        $this->assertFalse(file_exists($outputPath));

        $dataContent = file_get_contents(dirname(__FILE__).'/data.json');
        $data = json_decode($dataContent,true);

        //echo json_encode($data);

        //testing merge method
        $docxTemplate->merge($data,$outputPath,false,true);

        $this->assertTrue(file_exists($outputPath));

    }
}




?>
