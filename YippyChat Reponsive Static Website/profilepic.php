<?php
$uploaddir = '/home4/deano/etc/uploads/';
$name = $_GET['yname'];
$url = $uploaddir.'YacYo'.$name.'.jpg';
if(!file_exists($url))
$url = 'N.jpg';
$imginfo = getimagesize($url);
header("Content-type: " . $imginfo['mime']);
readfile($url);        
?>