<?php

// Boot the app
require_once __DIR__ . '/Autoloader.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;

// Create new presentation object
$obj = new PhpPresentation();


// BEGIN: Slide 01 -------------------------------------------------------------
// Get active slide - once object created a slide will be created
$slide = $obj->getActiveSlide();

// Create a shape (text)
$shape = $slide->createRichTextShape()
      ->setHeight(300)
      ->setWidth(600)
      ->setOffsetX(170)
      ->setOffsetY(180);
$textRun = $shape->createTextRun('Thank you for using PHPPresentation!');
$textRun->getFont()->setBold(true)
                   ->setSize(60);

// END: Slide 01

// BEGIN: Slide 02 -------------------------------------------------------------
// Create slide
$slide = $obj->createSlide();

// Create a shape (text)
$shape = $slide->createRichTextShape()
      ->setHeight(300)
      ->setWidth(600)
      ->setOffsetX(170)
      ->setOffsetY(180);
$textRun = $shape->createTextRun('Thank you for using PHPPresentation!');
$textRun->getFont()->setBold(true)
                   ->setSize(60);

// END: Slide 02

$oWriterPPTX = IOFactory::createWriter($obj, 'PowerPoint2007');
$oWriterPPTX->save(__DIR__ . "/sample.pptx");
