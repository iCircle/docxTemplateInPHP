<?php

include_once dirname(__FILE__).'/../../src/DocxTemplate.php'

class SimpleTest extends PHPUnit_Framework_TestCase{

    public function testTemplateCopy(){
        echo "Starting";
        $templatePath = dirname(__FILE__).'/template.docx';

        $docxTemplate = new DocxTemplate($templatePath);
        $docxTemplate->merge();

        echo "Success";
    }
}




?>
