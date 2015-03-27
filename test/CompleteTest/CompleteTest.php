<?php

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

        $data = array();
        $host = array();
        $host["name"] = "My Company";
        $host["address"] = "No25, Street 1, 2nd Avenue";
        $host["phoneNo"] = "+91 80 2236 1234";
        $host["emailId"] = "contact@mycompany.com";
        $host["tinNo"] = "021256369974";
        $host["panNo"] = "ABCDE1234Z";
        $host["cst"] = "CST1234546789";
        $host["logo"] = dirname(__FILE__).'/logo.png';
        $data["host"] = $host;

        $record = array();
        $record["no"] = "PO000001";
        $record["date"] = "26-Mar-2015";

        $quotation = array();
        $quotation["no"] = "QUOT000001";
        $quotation["date"] = "01-Mar-2015";
        $record["quotation"] = $quotation;

        $org = array();
        $org["name"] = "The Other Company";
        $org["address"] = "No65, Street 3, 1st Avenue";
        $record["org"] = $org;

        $record["netAmount"] = "25,000.00";

        $items = array();

        $item = array();
        $item["name"] = "P001";
        $item["quantity"] = "3";
        $item["rate"] = "900.00";
        $item["discount"] = "10";
        $item["totalPrice"] = "2,430.00";
        $items[] = $item;

        $item = array();
        $item["name"] = "P002";
        $item["quantity"] = "10";
        $item["rate"] = "1,500.00";
        $item["discount"] = "5";
        $item["totalPrice"] = "14,250.00";
        $items[] = $item;

        $item = array();
        $item["name"] = "P003";
        $item["quantity"] = "1";
        $item["rate"] = "8,320.00";
        $item["discount"] = "0";
        $item["totalPrice"] = "8,320.00";
        $items[] = $item;

        $record["items"] = $items;

        $data["record"] = $record;

        //testing merge method
        $docxTemplate->merge($data,$outputPath);

        $this->assertTrue(file_exists($outputPath));

    }
}




?>
