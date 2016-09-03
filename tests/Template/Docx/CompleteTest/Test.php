<?php

namespace icircle\tests\Template\Docx\CompleteTest;

use icircle\Template\Docx\DocxTemplate;
use icircle\tests\Template\Util;

class Test extends \PHPUnit_Framework_TestCase{

    public function testMerge(){
        $templatePath = dirname(__FILE__).'/template.docx';

        $docxTemplate = new DocxTemplate($templatePath);
        $outputPath = Util::createTempDir('/icircle/template/docx').'/mergedOutput.docx';
        
        $this->assertFalse(file_exists($outputPath));

        $dataContent = file_get_contents(dirname(__FILE__).'/data.json');
        $data = json_decode($dataContent,true);

        //testing merge method
        $docxTemplate->merge($data,$outputPath,false,true);

        $this->assertTrue(file_exists($outputPath));

    }
}




?>
