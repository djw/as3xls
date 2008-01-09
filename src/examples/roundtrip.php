<?php
$file = $_FILES['file'];

if(is_uploaded_file($file['tmp_name'])) {
	$filecont = file_get_contents($file['tmp_name']);
	echo(base64_encode($filecont));
} else {
	die();
} 
?>