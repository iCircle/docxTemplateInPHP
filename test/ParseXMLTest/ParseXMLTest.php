<?php

include_once dirname(__FILE__).'/../../src/DocxTemplate.php';

class ParseXMLTest extends PHPUnit_Framework_TestCase{

    public function testParseXML(){

        $templateDoc = '
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
	    <w:p w:rsidR="00927B0C" w:rsidRDefault="005E5E04">
		    <w:r>
			    <w:t>[host.na</w:t>
			</w:r>
			<w:r w:rsidRPr="005E5E04">
			    <w:t>me]</w:t>
			</w:r>
		</w:p>
		<w:p w:rsidR="005E5E04" w:rsidRDefault="005E5E04"/>
		<w:p w:rsidR="005E5E04" w:rsidRDefault="005E5E04" w:rsidP="005E5E04">
		    <w:pPr>
		        <w:pStyle w:val="Heading1"/>
			</w:pPr>
			<w:proofErr w:type="spellStart"/>
			<w:r>
			    <w:t>Address : [host.addre</w:t>
			    <w:t>ss</w:t>
			</w:r>
			<w:proofErr w:type="spellEnd"/>
	    </w:p>
		<w:p>
		    <w:r>
			    <w:t xml:space="preserve">Phone : [host.p</w:t>
			    <w:t>hon</w:t>
			    <w:t>e] </w:t>
			</w:r>
		</w:p>
		<w:p>
		    <w:r>
			    <w:t xml:space="preserve">Host No : [host.no]</w:t>
			</w:r>
		</w:p>
		<w:sectPr w:rsidR="005E5E04" w:rsidSect="00927B0C">
		    <w:pgSz w:w="12240" w:h="15840"/>
		    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
		    <w:cols w:space="720"/>
		    <w:docGrid w:linePitch="360"/>
		</w:sectPr>
	</w:body>
</w:document>';



        $document = new DOMDocument();
        $document->preserveWhiteSpace = false;
        $document->loadXML($templateDoc);

        $templatePath = dirname(__FILE__).'/template.docx';

        $template = new ZipArchive();
        $template->open($templatePath,ZipArchive::CREATE);
        $template->addFromString("word/document.xml",$document->saveXML());
        $template->close();

        $docxTemplate = new DocxTemplate($templatePath);
        $outputPath = dirname(__FILE__).'/mergedOutput.docx';
        if(file_exists($outputPath)){
            unlink($outputPath);
        }
        $this->assertFalse(file_exists($outputPath));

        //testing merge method
        $docxTemplate->merge(array("host"=>array("name"=>"My Company","no"=>"0")),$outputPath);

        $resultZip = new ZipArchive();
        $resultZip->open($outputPath);
        $resultZip->extractTo($outputPath."_");

        $resultDoc = new DOMDocument();
        $resultDoc->load($outputPath."_"."/word/document.xml");

        $expectedXML = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p w:rsidR="00927B0C" w:rsidRDefault="005E5E04"><w:r><w:t/></w:r><w:r w:rsidRPr="005E5E04"><w:t>My Company</w:t></w:r></w:p><w:p w:rsidR="005E5E04" w:rsidRDefault="005E5E04"/><w:p w:rsidR="005E5E04" w:rsidRDefault="005E5E04" w:rsidP="005E5E04"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:proofErr w:type="spellStart"/><w:r><w:t>Address : [host.addre</w:t><w:t>ss</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p><w:p><w:r><w:t xml:space="preserve">Phone : </w:t><w:t/><w:t>[host.phone] </w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Host No : 0</w:t></w:r></w:p><w:sectPr w:rsidR="005E5E04" w:rsidSect="00927B0C"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>';
        $this->assertTrue($resultDoc->saveXML($resultDoc->documentElement) == $expectedXML);

    }

}




?>
