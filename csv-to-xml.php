<?php
error_reporting(E_ALL | E_STRICT);
ini_set('display_errors', true);
ini_set('auto_detect_line_endings', true);

function PrepareXMLName($PrepareString)
	{
		
	$PrepareString = str_replace(" ","",$PrepareString);
	$PrepareString = preg_replace('#\W#', '', $PrepareString);
	$PrepareString = str_replace("ZSPAZESZ","",$PrepareString);
	$PrepareString = strtolower($PrepareString);
	
	return $PrepareString;
	}	

$inputFilename    = 'csv/1-2lVetPop11_POS_National.csv';
$outputFilename   = 'xml/2lVetPop11_POS_National.xml';

// Open csv to read
$inputFile  = fopen($inputFilename, 'rt');

// Get the headers of the file
$headers = fgetcsv($inputFile);

// Create a new dom document with pretty formatting
$doc  = new DomDocument();
$doc->formatOutput   = true;

// Add a root node to the document
$root = $doc->createElement('rows');
$root = $doc->appendChild($root);

// Loop through each row creating a <row> node with the correct data
while (($row = fgetcsv($inputFile)) !== FALSE)
	{
	$container = $doc->createElement('row');
	foreach ($headers as $i => $header)
		{
		$header = str_replace(chr(32),"_",trim($header));
		$header = strtolower($header);
		if($header==''){ $header = 'empty';}
		$header = PrepareXMLName($header);
		if(is_numeric($header)) { $header = "number-". $header; }
		//echo "HERE: " . $header . "<br />";  
		$child = $doc->createElement($header);
		$child = $container->appendChild($child);
		$value = $doc->createTextNode($row[$i]);
		$value = $child->appendChild($value);
		}
	$root->appendChild($container);
	}

header("Content-type: text/xml");
echo $doc->saveXML();