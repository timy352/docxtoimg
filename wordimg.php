<?php
class WordPHP
{
	private $file;
	private $rels_xml;
	private $doc_media = [];


	/**
	 * CONSTRUCTOR
	 * 
	 * @param Boolean $view View mode or not
	 * @return void
	 */
	public function __construct($view_=null)
	{
		if($view_ != null) {
			$this->view = $view_;
		}
	}

	/**
	 * READS The Relationships into a XML file
	 * 
	 * @param var $object The class variable to set as DOMDocument 
	 * @param var $xml The xml file
	 * @param string $encoding The encoding to be used
	 * @return void
	 */
	private function setXmlParts(&$object, $xml, $encoding)
	{
		$object = new DOMDocument();
		$object->encoding = $encoding;
		$object->preserveWhiteSpace = false;
		$object->formatOutput = true;
		$object->loadXML($xml);
		$object->saveXML();
	}


	/**
	 * READS The Relationships into a XML file
	 * 
	 * @param String $filename The filename
	 * @return void
	 */
	private function readZipPart($filename)
	{
		$zip = new ZipArchive();
		$_xml_rels = 'word/_rels/document.xml.rels';

		if (true === $zip->open($filename)) {
			//Get the relationships
			if (($index = $zip->locateName($_xml_rels)) !== false) {
				$xml_rels = $zip->getFromIndex($index);
			}
			$zip->close();
		} else die('non zip file');

		$enc = mb_detect_encoding($xml_rels);
		$this->setXmlParts($this->rels_xml, $xml_rels, $enc);
	}


	/**
	 * READS THE GIVEN DOCX FILE AND SAVES IMAGES
	 *  
	 * @param String $filename - The DOCX file name
	 * @return NONE
	 */
	public function readDocument($filename)
	{
		$Fdir = str_replace('.','_',$filename);
		if (!is_dir($Fdir)){
			mkdir($Fdir, 0755, true);
		}
		
		$this->readZipPart($filename);
		$reader = new XMLReader();
		$reader->XML($this->rels_xml->saveXML());
		
		while ($reader->read()) {
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name=='Relationship') {
				$Ftarget = $reader->getAttribute("Target");
				if (substr($Ftarget,0,11) == 'media/image'){
					$Fimg = "word/".$Ftarget;
					$zip = new ZipArchive();
					$im = null;
					if (true === $zip->open($filename)) {
						$image = $zip->getFromName($Fimg);
					
						$arr = explode('.', $Fimg);
						$l = count($arr);
						$ext = strtolower($arr[$l-1]);
						$Fimg = explode('/', $arr[$l-2]);
						$F = count($Fimg);
						$Iname = $Fimg[$F-1];
						$im = imagecreatefromstring($image);
						$fname = $Fdir.'/'.$Iname.'.'.$ext;
						switch ($ext) {
							case 'png':
								imagepng($im, $fname);
								break;
							case 'bmp':
								imagebmp($im, $fname);
								break;
							case 'gif':
								imagegif($im, $fname);
								break;
							case 'jpeg':
							case 'jpg':
								imagejpeg($im, $fname);
								break;
							default:
								break;
						}
					}
					imagedestroy($im);
					if ($this->view){
						echo $fname."<br>";
					}
					$zip->close();
				}
			}
		}
	}
}
	
					


