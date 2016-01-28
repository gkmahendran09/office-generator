<?php
ini_set('error_reporting', E_ALL);
ini_set('display_errors', 'On');  //On or Off

date_default_timezone_set("Asia/Calcutta");

// Autoload Libraries
require 'vendor/autoload.php';

use Templates\PowerPoint\Template1 as PowerPointTemp1;
use Templates\Word\Template1 as WordTemp1;

/*
 * Generate PowerPoint
 *
 */
$obj = new PowerPointTemp1(); // Create presentation object
$obj->removeSlideByIndex(0); // Remove First Slide

for( $i=1; $i<=2; $i++ ) {
  // Create Slide
  $PowerPoint = $obj->create(file_get_contents("./data/data{$i}.json"));
}

$obj->save($PowerPoint, __DIR__ . "/output/sample.pptx"); // Save Presentation

/* -----------------------------------------------------------------------------

/*
 * Generate Word
 *
 */
$obj = new WordTemp1(); // Creating the new document...

for( $i=1; $i<=2; $i++ ) {
   // Create Page
  $Word = $obj->create(file_get_contents("./data/data{$i}.json"));
}

$obj->save($Word, __DIR__ . "/output/sample.docx"); // Save Word

?>

<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- The above 3 meta tags *must* come first in the head; any other head content must come *after* these tags -->
    <meta name="description" content="">
    <meta name="author" content="">
    <link rel="icon" href="../../favicon.ico">

    <title>Office Generator</title>

    <!-- Bootstrap core CSS -->
    <link href="assets/css/bootstrap.min.css" rel="stylesheet">
    <link href="assets/css/jumbotron-narrow.css" rel="stylesheet">

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
  </head>

  <body>

    <div class="container">
      <div class="header clearfix">
        <h3 class="text-muted">Office Generator</h3>
      </div>

      <div class="jumbotron">
        <h1>Dynamic<br>PPTX &amp; DOCX</h1>
        <p>
          <a class="btn btn-lg btn-success" href='<?php echo "output/sample.pptx"; ?>' role="button">.pptx</a>
          <a class="btn btn-lg btn-danger" href='<?php echo "output/sample.docx"; ?>' role="button">.docx</a>
        </p>
      </div>
    </div> <!-- /container -->

  </body>
</html>
