<?php

namespace icircle\tests\Template\Docx\HyperlinkTest;

use icircle\Template\Docx\DocxTemplate;
use icircle\tests\Template\Util;

class Test extends \PHPUnit_Framework_TestCase{

    public function testHyperlink(){
        
        $tempDir = Util::createTempDir('/icircle/template/docx');

        $templatePath = $tempDir.'/template.docx';
        copy(dirname(__FILE__).'/template.docx', $templatePath);

        $docxTemplate = new DocxTemplate($templatePath);
        $outputPath = $tempDir.'/mergedOutput.docx';
        
        $this->assertFalse(file_exists($outputPath));

        //testing merge method
        $data = array(
                "link1"=>array(
                        "text"=>"LINK1_TEXT",
                        "st"=>"LINK1_ST",
                        "target"=>"LINK1_TARGET"
                ),
                "link2"=>array(
                        "text"=>"LINK2_TEXT",
                        "st"=>"LINK2_ST",
                        "target"=>"LINK2_TARGET"
                ),
                "link3"=>array(
                        "text"=>"LINK3_TEXT",
                        "st"=>"LINK3_ST",
                        "target"=>"LINK3_TARGET"
                ),
                "link4"=>array(
                        "text"=>"LINK4_TEXT",
                        "st"=>"LINK4_ST",
                        "target"=>"LINK4_TARGET",
                        "targetSubject"=>"LINK4_TARGET_SUBJECT"
                )
        );
        $docxTemplate->merge($data,$outputPath);

        $resultZip = new \ZipArchive();
        $resultZip->open($outputPath);
        $resultZip->extractTo($outputPath."_");

        $resultDoc = new \DOMDocument();
        $resultDoc->load($outputPath."_"."/word/document.xml");
        $expectedDocumentXML = '<w:document xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:body><w:p w:rsidR="0005415C" w:rsidRDefault="00F97E66"><w:bookmarkStart w:id="0" w:name="_top"/><w:bookmarkEnd w:id="0"/><w:r><w:t xml:space="preserve">line1Start </w:t></w:r><w:hyperlink r:id="rId4" w:tooltip="hyperlink1STLINK1_ST" w:history="1"><w:r w:rsidR="00E76BD6"><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>hyperLINK1_TEXTlink1</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> line1End</w:t></w:r></w:p><w:p w:rsidR="00B35265" w:rsidRDefault="00F97E66"><w:r><w:t xml:space="preserve">line2Start </w:t></w:r><w:hyperlink w:anchor="_top" w:tgtFrame="_blank" w:tooltip="LINK2_SThyperlink2ST" w:history="1"><w:r w:rsidR="00E76BD6"><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>hyperLINK2_TEXT link2</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> line2End</w:t></w:r></w:p><w:p w:rsidR="00B35265" w:rsidRDefault="00F97E66"><w:r><w:t xml:space="preserve">line3Start </w:t></w:r><w:hyperlink r:id="rId5" w:tooltip="hyperlink3LINK3_STST" w:history="1"><w:r w:rsidR="00E76BD6"><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>hyper LINK3_TEXTlink3</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> line3End</w:t></w:r></w:p><w:p w:rsidR="00B35265" w:rsidRDefault="00F97E66"><w:r><w:t xml:space="preserve">line4Start </w:t></w:r><w:hyperlink r:id="rId6" w:tooltip="hyperlink4ST LINK4_ST" w:history="1"><w:r w:rsidR="00E76BD6"><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>hyper LINK4_TEXT link4</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> line4End</w:t></w:r></w:p><w:sectPr w:rsidR="00B35265" w:rsidSect="0005415C"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>';
        $this->assertTrue($resultDoc->saveXML($resultDoc->documentElement) == $expectedDocumentXML);
        
        $resultDoc = new \DOMDocument();
        $resultDoc->load($outputPath."_"."/word/_rels/document.xml.rels");
        $expectedDocumentXML = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:LINK4_TARGET?subject=hyperlink4emailSubjectLINK4_TARGET_SUBJECT" TargetMode="External"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="file:///C:\Users\Raaghu\Desktop\newDocLINK3_TARGET.docx" TargetMode="External"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="LINK1_TARGET.html" TargetMode="External"/></Relationships>';
        $this->assertTrue($resultDoc->saveXML($resultDoc->documentElement) == $expectedDocumentXML);
    }

}


?>
