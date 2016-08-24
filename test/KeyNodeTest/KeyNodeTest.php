<?php

use icircle\Template\KeyNode;

class KeyNodeTest extends PHPUnit_Framework_TestCase{

    public function testKeyNode(){
        $domDocument = new DOMDocument();
        $wtElement = $domDocument->createElement("w:t");

        $key = "[item;repeat=record.items;repeatType=row]";

        $keyNode = new KeyNode($key,true,true,$wtElement);

        $this->assertTrue($keyNode->key() == "item");
        $this->assertTrue($keyNode->originalKey() == $key);
        $options = $keyNode->options();
        $this->assertTrue(count($options) == 2);
        $this->assertTrue($options["repeat"] == "record.items");
        $this->assertTrue($options["repeatType"] == "row");

        $keyNode = new KeyNode($key,false,true,$wtElement);

        $this->assertTrue($keyNode->key() == "[item;repeat=record.items;repeatType=row]");
        $options = $keyNode->options();
        $this->assertTrue(count($options) == 0);

    }
}




?>
