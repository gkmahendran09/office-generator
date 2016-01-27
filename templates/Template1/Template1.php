<?php
namespace Template1;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\IOFactory;

class Template1
{

  public static function createTextRun($shape, $text, $size, $color, $isBold=false)
  {
    $textRun = $shape->createTextRun($text);
    $textRun->getFont()
            ->setSize($size)
            ->setColor(new Color($color));

    if($isBold) {
      $textRun->getFont()->setBold(true);
    }
  }

  public static function createRow($shape)
  {
    $row = $shape->createRow();
    $row->setHeight(50);
    return $row;
  }

  public static function createCell($cell, $text)
  {
    $cell->getBorders()->getBottom()->setColor(new Color(F2F2F2));
    $cell->getBorders()->getTop()->setColor(new Color(F2F2F2));
    $cell->getBorders()->getLeft()->setColor(new Color(F2F2F2));
    $cell->getBorders()->getRight()->setColor(new Color(F2F2F2));
    $cell->createTextRun($text)->getFont()->setSize(14)->setColor( new Color( '000000' ) );
  }

  public static function createDrawingShape($currentSlide, $name, $description, $path, $width, $height, $offsetX, $offsetY)
  {
    $shape = new Drawing();
    // $shape = $currentSlide->createDrawingShape();
    $shape->setName($name)
          ->setDescription($description)
          ->setPath($path)
          // ->setWidth($width)
          ->setHeight($height)
          ->setOffsetX($offsetX)
          ->setOffsetY($offsetY);
    $currentSlide->addShape($shape);
  }

  public static function createLineShape($currentSlide, $fromX, $fromY, $toX, $toY)
  {
    $line = new Line($fromX, $fromY, $toX, $toY);
    $currentSlide->addShape($line);
  }

  public static function createRichTextShape($currentSlide, $width, $height, $offsetX, $offsetY)
  {
    $shape = $currentSlide->createRichTextShape()
         ->setHeight($height)
         ->setWidth($width)
         ->setOffsetX($offsetX)
         ->setOffsetY($offsetY);

    $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
                                                ->setMarginLeft(25);
    return $shape;

  }

}
