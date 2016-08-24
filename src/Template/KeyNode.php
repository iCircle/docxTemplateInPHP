<?php

namespace icircle\Template;

class KeyNode {

	private $key = null;
	private $originalKey = null;
	private $isKey = false;
	private $isComplete = false;
	private $element = null;
	private $options = array();
	
	function __construct($key,$isKey,$isComplete,DOMElement $element){
		$this->key = $key;
		$this->originalKey = $key;
		$this->isKey = $isKey;
		$this->isComplete = $isComplete;
		$this->element = $element;
	
		if($this->isKey && $this->isComplete){
			//parse the complete key to extract options
			$options = preg_split("/;/",substr($this->key,1,strlen($this->key)-2));
			if(count($options) > 1){
				for($i=1;$i<count($options);$i++){
					$option = preg_split("/=/",$options[$i]);
					$optionName = trim($option[0]);
					$optionValue = true;
					if(count($option) > 1){
						$optionValue = trim($option[1]);
					}
					$this->options[$optionName] = $optionValue;
				}
			}
			$this->key = trim($options[0]);
			//echo "\n".$this->key;
		}
	
	}
	
	/**
	 * @return string
	 */
	public function key(){
		return $this->key;
	}
	
	/**
	 * @return string
	 */
	public function originalKey(){
		return $this->originalKey;
	}
	
	/**
	 * @return int
	 */
	public function isKey(){
		return $this->isKey;
	}
	
	/**
	 * @return boolean
	 */
	public function isComplete(){
		return $this->isComplete;
	}
	
	/**
	 * @return DOMElement
	 */
	public function element(){
		return $this->element;
	}
	
	/**
	 * @return array
	 */
	public function options(){
		return $this->options;
	}
	
	
}