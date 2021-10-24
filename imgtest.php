<!DOCTYPE html>
<html lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />

</head>

<body>
<?php
require_once('wordimg.php');
$rt = new WordPHP(true);
$rt->readDocument('sample.docx');
?>
</body>
