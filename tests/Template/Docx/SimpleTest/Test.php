<?php

namespace icircle\tests\Template\Docx\SimpleTest;

use icircle\Template\Docx\DocxTemplate;
use icircle\tests\Template\Util;

class Test extends \PHPUnit_Framework_TestCase{

    public function testMerge(){
        $templatePath = dirname(__FILE__).'/template.docx';

        $docxTemplate = new DocxTemplate($templatePath);
        $outputPath = Util::createTempDir('/icircle/template/docx').'/mergedOutput.docx';

        $this->assertFalse(file_exists($outputPath));
        //testing merge method
        $docxTemplate->merge(array(),$outputPath);
        $this->assertTrue(file_exists($outputPath));

    }
}




?>
