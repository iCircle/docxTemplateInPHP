# docxTemplateInPHP
Create Templates in MS Word docx format and use them Creating Business documents using PHP 

## Introduction
This library merges the data in php array into docx file.

The template docx file can be created in MS Word, 
To specify keys use [key] format in MS Word (i.e, specify the array indices inside the square bracket)

### Example
for the data 
```json
{
    "host":{
        "name":"Host Company",
        "address":"1st Main, 2nd Cross, Bangalore"
    }
}
```

the possible keys are
[host.name]
[host.address]

## Basic Usage
```php
include_once 'DocxTemplate.php';

$docxTemplate = new DocxTemplate('path/to/template/file');
$dataArray = array() // fill the $dataArray with data
$docxTemplate->merge($dataArray,'path/to/output/file');
```

For More Usage please see the included tests or contact me
