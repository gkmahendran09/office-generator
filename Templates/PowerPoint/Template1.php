<?php
namespace Templates\PowerPoint;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;

use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Shape\Line;


class Template1 extends PhpPresentation implements \Templates\TemplateInterface
{
    private function createTextRun($shape, $text, $size, $color, $isBold=false)
    {
      $textRun = $shape->createTextRun($text);
      $textRun->getFont()
              ->setSize($size)
              ->setColor(new Color($color));

      if($isBold) {
        $textRun->getFont()->setBold(true);
      }
    }

    private function createRow($shape)
    {
      $row = $shape->createRow();
      $row->setHeight(50);
      return $row;
    }

    private function createCell($cell, $text)
    {
      $cell->getBorders()->getBottom()->setColor(new Color(F2F2F2));
      $cell->getBorders()->getTop()->setColor(new Color(F2F2F2));
      $cell->getBorders()->getLeft()->setColor(new Color(F2F2F2));
      $cell->getBorders()->getRight()->setColor(new Color(F2F2F2));
      $cell->createTextRun($text)->getFont()->setSize(14)->setColor( new Color( '000000' ) );
    }

    private function createDrawingShape($currentSlide, $name, $description, $path, $width, $height, $offsetX, $offsetY)
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

    private function createLineShape($currentSlide, $fromX, $fromY, $toX, $toY)
    {
      $line = new Line($fromX, $fromY, $toX, $toY);
      $currentSlide->addShape($line);
    }

    private function createRichTextShape($currentSlide, $width, $height, $offsetX, $offsetY)
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

    public function create($json)
    {
      $data = json_decode(utf8_encode($json));

      // set properties
      $this->getProperties()->setCreator($data->properties->Creator)
                            ->setLastModifiedBy($data->properties->LastModifiedBy)
                            ->setTitle($data->properties->Title)
                            ->setSubject($data->properties->Subject)
                            ->setDescription($data->properties->Description)
                            ->setKeywords($data->properties->Keywords)
                            ->setCategory($data->properties->Category);


      // Create Slide
      $currentSlide = $this->createSlide();

      // Title
      $shape = $currentSlide->createRichTextShape()
           ->setHeight(140)
           ->setWidth(960)
           ->setOffsetX(0)
           ->setOffsetY(0)
           ->setInsetTop(40);
      $shape->getFill()->setFillType(Fill::FILL_SOLID)
                    ->setRotation(90)
                    ->setStartColor(new Color('BC141A'))
                    ->setEndColor(new Color('BC141A'));

      $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
                                                  ->setMarginLeft(25);

      $textRun = $shape->createTextRun($data->content->title);
      $textRun->getFont()
                      ->setSize(28)
                      ->setColor( new Color( 'FFFFFF' ) );

      // Flag
      $shape = new Drawing();
      $shape->setName($data->flag->name)
            ->setDescription($data->flag->description)
            ->setPath($data->flag->path)
            ->setHeight(85)
            ->setOffsetX(850)
            ->setOffsetY(27);
      $currentSlide->addShape($shape);

      // Client name
      $shape = $currentSlide->createRichTextShape()
           ->setHeight(40)
           ->setWidth(960)
           ->setOffsetX(0)
           ->setOffsetY(140)
           ->setInsetTop(15);
      $shape->getFill()->setFillType(Fill::FILL_SOLID)
                    ->setRotation(90)
                    ->setStartColor(new Color('FFFFFF'))
                    ->setEndColor(new Color('FFFFFF'));

      $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
                                                  ->setMarginLeft(25);

      $textRun = $shape->createTextRun($data->content->client);
      $textRun->getFont()
                      ->setSize(18)
                      ->setColor( new Color( '000000' ) );


      // Services
      $shape = $currentSlide->createRichTextShape()
           ->setHeight(450)
           ->setWidth(960)
           ->setOffsetX(0)
           ->setOffsetY(200);
      $shape->getFill()->setFillType(Fill::FILL_SOLID)
                    ->setRotation(90)
                    ->setStartColor(new Color('F2F2F2'))
                    ->setEndColor(new Color('F2F2F2'));

      $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
                                                  ->setMarginLeft(25);

      $textRun = $shape->createTextRun($data->content->subheading);
      $textRun->getFont()
              ->setBold(true)
              ->setSize(16)
              ->setColor( new Color( '000000' ) );

      $shape = $currentSlide->createRichTextShape()
           ->setWidth(910)
           ->setOffsetX(50)
           ->setOffsetY(220);


           //ROW 3
           $shape = $currentSlide->createTableShape(1);
           $shape->setHeight(140);
           $shape->setWidth(800);
           $shape->setOffsetX(75);
           $shape->setOffsetY(250);

           foreach($data->list as $key => $value) {
             $row = $this->createRow($shape);
             $cell = $row->nextCell();
             $this->createCell($cell, $value);
           }

           $offsetY = 250;
           foreach($data->iconPath as $key => $value) {
             $this->createDrawingShape($currentSlide, 'Icon', 'Icon', $value, 35, 34, 40, $offsetY);
             $offsetY += 50;
           }

          //  createDrawingShape($currentSlide, $name, $description, $path, $width, $height, $offsetX, $offsetY)
           $this->createDrawingShape($currentSlide, $data->mainImage->name, $data->mainImage->description, $data->mainImage->path, 430, 300, 444, 200);

          //  $fromX, $fromY, $toX, $toY
           $this->createLineShape($currentSlide, 0, 500, 960, 500);


           $shape = $this->createRichTextShape($currentSlide, 960, 150, 0, 500);
           $this->createTextRun($shape, $data->content->contentTitle, "18", "000000", true);
           $shape->createBreak();
           $this->createTextRun($shape, $data->content->contentDescription, "18", "000000");

           $path = "./resources/slide/image1.png";
           $width = 100;
           $height = 50;
           $offsetX = 25;
           $offsetY = 660;
           $this->createDrawingShape($currentSlide, 'JLL Logo', 'JLL Logo', $path, $width, $height, $offsetX, $offsetY);

           $shape = $this->createRichTextShape($currentSlide, 10, 15, 900, 675);
           $this->createTextRun($shape, "1", "11", "AAAAAA");


      return $this;
    }

    public function save($obj, $path)
    {
      $oWriterPPTX = IOFactory::createWriter($obj, 'PowerPoint2007');
      $oWriterPPTX->save($path);
    }
}
