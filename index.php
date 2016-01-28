<?php

// Autoload Libraries
require 'vendor/autoload.php';

use Templates\PowerPoint\Template1 as PowerPointTemp1;

// Create presentation object
$obj = new PowerPointTemp1();

// Remove First Slide
$obj->removeSlideByIndex(0);

for( $i=1; $i<=2; $i++ ) {
  // Create Slide
  $PowerPoint = $obj->create(file_get_contents("./data/data{$i}.json"));
}

// Save Presentation
$obj->save($PowerPoint, __DIR__ . "/sample.pptx");