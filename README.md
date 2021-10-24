# A php class to extract all the images from a Word DOCX document and save them as separate image files

## Description

This php class will take a DOCX type Word document and extract all the images that are contined in it. They will be then all be saved in a directory with the same name as the original DOCX file. This directory will be automatically created if it does not exist. In the normal mode this class will not provide any output to screen. A php demonstration file (imgtest.php) is included. This demonstration file expects the DOCX file to be named 'sample.docx'.

# USAGE

## Normal mode to save all the images (no output to screen)
```
$rt = new WordPHP(false); or $rt = new WordPHP();
```

## Display on screen the names of the images found and saved
```
$rt = new WordPHP(true);
```

## Read docx file and save all the images found
```
$rt->readDocument('FILENAME');
```
