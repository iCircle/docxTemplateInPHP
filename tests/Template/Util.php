<?php

namespace icircle\tests\Template;

class Util {
	public static function createTempDir($dir = null){
		$tempDir = sys_get_temp_dir();
		if($dir != null){
			$tempDir .= "/".$dir;
		}
		$tempDir .= "/".microtime(false);
		
		mkdir($tempDir,null,true);
		
		return $tempDir;
	}
}