<?php
/**
 * User: Raghavendra K R
 * Date: 16/3/15
 * Time: 10:05 AM
 */

class DocxTemplate {
    private $template = null;

    function __construct($templatePath){
        if(!file_exists($templatePath)){
            throw new Exception("Invalid Template Path");
        }
        $this->template = $templatePath;
    }

    function merge($data){
        //open the Archieve to a temp folder

        $workingFile = tempnam(sys_get_temp_dir(),'');
        if($workingFile === FALSE || !copy($this->template,$workingFile)){
            throw new Exception("Error in initializing working copy of the template");
        }




    }



}
