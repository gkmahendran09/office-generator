<?php

// Boot the app
require_once __DIR__ . '/boot.php';

// Include the template
require_once __DIR__ . '/templates/Template1/Template1.php';

// Use the template
use Templates\Template1;

use PhpOffice\PhpPresentation\IOFactory;


$obj = new Template1();

// Get active slide
$slide = $obj->createFirstSlide();

// Create a shape (text)
$shape = $slide->createRichTextShape()
      ->setHeight(300)
      ->setWidth(600)
      ->setOffsetX(170)
      ->setOffsetY(180);
$textRun = $shape->createTextRun('Thank you for using PHPPresentation!');
$textRun->getFont()->setBold(true)
                   ->setSize(60);


// Create slide
$slide = $obj->createNextSlide();

// Create a shape (text)
$shape = $slide->createRichTextShape()
      ->setHeight(300)
      ->setWidth(600)
      ->setOffsetX(170)
      ->setOffsetY(180);
$textRun = $shape->createTextRun('Thank you for using PHPPresentation!');
$textRun->getFont()->setBold(true)
                   ->setSize(60);

$oWriterPPTX = IOFactory::createWriter($obj, 'PowerPoint2007');
$oWriterPPTX->save(__DIR__ . "/sample.pptx");
