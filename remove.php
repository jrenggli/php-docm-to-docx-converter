<?php
/***************************************************************
*  Copyright notice
*
*  (c) 2011 Jonas Renggli (jonas.renggli@entris-banking.ch)
*  All rights reserved
*
*  This script is free software; you can redistribute it and/or modify
*  it under the terms of the GNU General Public License as published by
*  the Free Software Foundation; either version 2 of the License, or
*  (at your option) any later version.
*
*  The GNU General Public License can be found at
*  http://www.gnu.org/copyleft/gpl.html.
*  A copy is found in the textfile GPL.txt and important notices to the license
*  from the author is found in LICENSE.txt distributed with these scripts.
*
*
*  This script is distributed in the hope that it will be useful,
*  but WITHOUT ANY WARRANTY; without even the implied warranty of
*  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
*  GNU General Public License for more details.
*
*  This copyright notice MUST APPEAR in all copies of the script!
***************************************************************/

$filesToDelete = array();

$path = sys_get_temp_dir() . DIRECTORY_SEPARATOR . uniqid() . DIRECTORY_SEPARATOR;

/*
 * extract
 */
$zip = new ZipArchive;
if ($zip->open('input.docm') === TRUE) {
	$zip->extractTo($path);
	$zip->close();
} else {
}
unset($zip);




/*
 * delete macros
 * Credits: http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2008/01/11/2608.aspx
 */



//TODO: [Content_types.xml] bearbeiten
$filename = $path . '[Content_Types].xml';
$xml = simplexml_load_file($filename);
$ns = $xml->getNamespaces();
$xml->registerXPathNamespace('x', $ns[""]);

$xpath = $xml->xpath('//x:Override[@PartName="/word/document.xml"]');
$xpath[0]['ContentType'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

$xpath = $xml->xpath('//x:Override[@PartName="/word/vbaData.xml"]');
$dom = dom_import_simplexml($xpath[0]);
$dom->parentNode->removeChild($dom);
$filesToDelete[] = 'word/vbaData.xml';

$xpath = $xml->xpath('//x:Default[@Extension="bin"]');
$dom = dom_import_simplexml($xpath[0]);
$dom->parentNode->removeChild($dom);

$xml->asXML($filename);

//TODO: document.xml.rels bearbeiten
$filename = $path . 'word/_rels/document.xml.rels';
$xml = simplexml_load_file($filename);

$ns = $xml->getNamespaces();
$xml->registerXPathNamespace('x', $ns[""]);

$xpath = $xml->xpath('//x:Relationship[@Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"]');

$target = (string) $xpath[0]['Target'];
$filesToDelete[] = 'word/' . $target;
$filesToDelete[] = 'word/_rels/' . $target . '.rels';

$dom = dom_import_simplexml($xpath[0]);
$dom->parentNode->removeChild($dom);

$xml->asXML($filename);

//TODO: vbaData.xml löschen

//TODO: vbaProject.bin löschen

//TODO: vbaProject.bin.rels löschen

foreach ($filesToDelete as $filename) {
    unlink($path . $filename);
}


/*
 * zip content
 * Credits: http://www.webandblog.com/hacks/zip-a-folder-on-the-server-with-php/
 */
$zip = new ZipArchive();
if ($zip->open("output.docx", ZIPARCHIVE::CREATE) !== TRUE) {
	die ("Could not open archive");
}

$iterator = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($path));
foreach ($iterator as $key => $value) {
	$localPath = str_replace($path, '', $key);
	if (!is_file(realpath($key))) {
		continue;
	}
	echo($localPath . '<br>');
	$zip->addFile(realpath($key), $localPath) or die ("ERROR: Could not add file: $key");
}

$zip->close();
echo "Archive created successfully.";
unset($zip);

//TODO: temporären Ordner löschen




?>
