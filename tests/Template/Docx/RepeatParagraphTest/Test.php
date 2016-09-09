<?php

namespace icircle\tests\Template\Docx\RepeatParagraphTest;

use icircle\Template\Docx\DocxTemplate;
use icircle\tests\Template\Util;

class Test extends \PHPUnit_Framework_TestCase{

    public function testTextRepeating(){

        $document = new \DOMDocument();
        $document->preserveWhiteSpace = false;
        $document->load(dirname(__FILE__).'/document.xml');
        
        $tempDir = Util::createTempDir('/icircle/template/docx');

        $templatePath = $tempDir.'/template.docx';
        copy(dirname(__FILE__).'/template.docx', $templatePath);

        $template = new \ZipArchive();
        $template->open($templatePath,\ZipArchive::CREATE);
        $template->addFromString("word/document.xml",$document->saveXML());
        $template->close();

        $docxTemplate = new DocxTemplate($templatePath);
        $outputPath = Util::createTempDir('/icircle/template/docx').'/mergedOutput.docx';
        
        $this->assertFalse(file_exists($outputPath));

        //testing merge method
        $data = array("host"=>array("name"=>"My Company"));
        $record = array("items"=>array(array("name"=>"item1"),array("name"=>"item2"),array("name"=>"item3")));
        $data["record"] = $record;
        $docxTemplate->merge($data,$outputPath);

        $resultZip = new \ZipArchive();
        $resultZip->open($outputPath);
        $resultZip->extractTo($outputPath."_");

        $resultDoc = new \DOMDocument();
        $resultDoc->load($outputPath."_"."/word/document.xml");

        $expectedXML = '<w:document xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body><w:p w:rsidR="00B65A93" w:rsidRDefault="00AB4A00"><w:r><w:t>Paragraph 1</w:t></w:r></w:p><w:p w:rsidP="00AD263B" w:rsidR="00AD263B" w:rsidRDefault="00AD263B"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Paragraph 2</w:t></w:r></w:p><w:p w:rsidP="00E70DA1" w:rsidR="00E70DA1" w:rsidRDefault="00E70DA1"><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:cs="Arial" w:hAnsi="Arial"/><w:b/><w:i/><w:color w:val="FF0000"/><w:sz w:val="40"/><w:u w:val="single"/></w:rPr></w:pPr><w:r w:rsidRPr="00E70DA1"><w:rPr><w:rFonts w:ascii="Arial" w:cs="Arial" w:hAnsi="Arial"/><w:b/><w:i/><w:color w:val="FF0000"/><w:sz w:val="40"/><w:u w:val="single"/></w:rPr><w:t>Paragraph 3</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:type="auto" w:w="0"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="2310"/><w:gridCol w:w="2310"/><w:gridCol w:w="2311"/><w:gridCol w:w="2311"/></w:tblGrid><w:tr w:rsidR="002165C9" w:rsidTr="0058431E"><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2310"/></w:tcPr><w:p w:rsidP="002165C9" w:rsidR="002165C9" w:rsidRDefault="006524DD"><w:r><w:t>Table Header1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2310"/></w:tcPr><w:p w:rsidP="002165C9" w:rsidR="002165C9" w:rsidRDefault="006524DD"><w:r><w:t>Table Header2</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2311"/></w:tcPr><w:p w:rsidP="002165C9" w:rsidR="002165C9" w:rsidRDefault="006524DD"><w:r><w:t>Table Header3</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2311"/></w:tcPr><w:p w:rsidP="002165C9" w:rsidR="002165C9" w:rsidRDefault="006524DD"><w:r><w:t>Table Header4</w:t></w:r></w:p></w:tc></w:tr><w:tr w:rsidR="006524DD" w:rsidTr="0058431E"><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2310"/></w:tcPr><w:p w:rsidP="006524DD" w:rsidR="006524DD" w:rsidRDefault="006524DD"><w:r><w:t>item1Cell 11</w:t></w:r></w:p><w:p w:rsidP="006524DD" w:rsidR="006524DD" w:rsidRDefault="006524DD"><w:r><w:t>item2Cell 11</w:t></w:r></w:p><w:p w:rsidP="006524DD" w:rsidR="006524DD" w:rsidRDefault="006524DD"><w:r><w:t>item3Cell 11</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2310"/></w:tcPr><w:p w:rsidP="006524DD" w:rsidR="006524DD" w:rsidRDefault="006524DD"><w:r><w:t/></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2311"/></w:tcPr><w:p w:rsidP="00912DA4" w:rsidR="006524DD" w:rsidRDefault="006524DD"><w:r><w:t>Cell13</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2311"/></w:tcPr><w:p w:rsidP="00912DA4" w:rsidR="006524DD" w:rsidRDefault="006524DD"><w:r><w:t>Cell14</w:t></w:r></w:p></w:tc></w:tr><w:tr w:rsidR="006524DD" w:rsidTr="0058431E"><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2310"/></w:tcPr><w:p w:rsidP="00912DA4" w:rsidR="006524DD" w:rsidRDefault="006524DD"><w:r><w:t>Cell 2</w:t></w:r><w:r><w:t>1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2310"/></w:tcPr><w:p w:rsidP="00912DA4" w:rsidR="006524DD" w:rsidRDefault="006524DD"><w:r><w:t>Cell2</w:t></w:r><w:r><w:t>2</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2311"/></w:tcPr><w:p w:rsidP="00912DA4" w:rsidR="006524DD" w:rsidRDefault="006524DD"><w:r><w:t>Cell2</w:t></w:r><w:r><w:t>3</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2311"/></w:tcPr><w:p w:rsidP="00912DA4" w:rsidR="006524DD" w:rsidRDefault="006524DD"><w:r><w:t>Cell2</w:t></w:r><w:r><w:t>4</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w:rsidP="002165C9" w:rsidR="002165C9" w:rsidRDefault="002165C9" w:rsidRPr="00E70DA1"/><w:sectPr w:rsidR="002165C9" w:rsidRPr="00E70DA1" w:rsidSect="00546BAC"><w:headerReference r:id="rId6" w:type="default"/><w:pgSz w:h="16838" w:w="11906"/><w:pgMar w:bottom="1440" w:footer="708" w:gutter="0" w:header="708" w:left="1440" w:right="1440" w:top="1440"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>';
        
        $this->assertTrue($resultDoc->saveXML($resultDoc->documentElement) == $expectedXML);
    }

}


?>
