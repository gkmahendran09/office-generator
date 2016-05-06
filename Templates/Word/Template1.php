<?php
namespace Templates\Word;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;



class Template1 extends PhpWord implements \Templates\TemplateInterface
{
    public function create($json)
    {
      $data = json_decode(utf8_encode($json));

      $section = $this->addSection();


      // $styleTable = array('borderSize' => 6, 'borderColor' => '999999');
      $styleTable = array(
                    'cellMarginTop' => 60,
                    'cellMarginBottom' => 60,
                    'cellMarginLeft' => 200,
                    'cellMarginRight' => 200,
                    // 'borderSize' => 6, 'borderColor' => '999999'
                  );
      $cellHCentered = array('align' => 'center');
      $cellVCentered = array('valign' => 'center');
      $styleCell = array('valign' => 'center');
      $styleClientSection = array('gridSpan' => 2, 'valign' => 'center');
      $styleMainSection = array('gridSpan' => 3, 'valign' => 'center', 'bgColor' => 'F2F2F2');

      $styleFirstRow = array('bgColor' => 'BC141A');
      $pageNumberFontStyle = array('size' => 10, 'color' => '999999', 'align' => 'right');
      $headerFontStyle = array('size' => 18, 'color' => 'ffffff', 'align' => 'left');
      $clientFontStyle = array('size' => 12, 'color' => '000000', 'align' => 'left');
      $subheadingFontStyle = array('size' => 12, 'bold' => true, 'color' => '000000', 'align' => 'left');
      $descriptionFontStyle = array('size' => 12, 'color' => '000000', 'align' => 'left');
      $listFontStyle = array('size' => 12, 'color' => '000000', 'align' => 'right');
      $imageStyle = array('align' => 'right');

      $this->addTableStyle('Header', $styleTable, $styleFirstRow);
      $table = $section->addTable('Header');

      // Row 1
      $table->addRow(800);
      $table->addCell(10000, $styleCell)
            ->addText(htmlspecialchars($data->content->title), $headerFontStyle);
      $table->addCell(2000, $styleCell)
            ->addImage($data->flag->path, $imageStyle);

      // Row 2
      $table->addRow(100);
      $table->addCell(12000, $styleClientSection)
            ->addText(htmlspecialchars($data->content->client), $clientFontStyle);


      $this->addTableStyle('test', array(
                    'cellMarginTop' => 60,
                    'cellMarginBottom' => 60,
                    'cellMarginLeft' => 200,
                    'cellMarginRight' => 200,
                    // 'borderSize' => 6, 'borderColor' => '999999'
                  ));
      $table = $section->addTable('test');

      $table->addRow(100);
      $listCell = $table->addCell(12000,  array('valign' => 'top', 'bgColor' => 'F2F2F2', 'borderBottomColor' => '999999', 'borderBottomSize' => '1'));
      $listCell->addText(htmlspecialchars($data->content->subheading), $subheadingFontStyle);
      foreach($data->list as $key => $value) {
        // $listCell->addTextBreak(1);

        $listCell->addImage($data->iconPath->{$key}, array(
          'width'         => 35,
          'height'        => 35,
          'positioning'      => 'relative'
        ));
        $listCell->addText(htmlspecialchars($value), $listFontStyle);
        // $cell->addImage($data->iconPath->{$key});
      }

      $cell = $table->addCell(4000,  array('valign' => 'center', 'bgColor' => 'F2F2F2', 'borderBottomColor' => '999999', 'borderBottomSize' => '1'));

      $cell->addImage($data->mainImage->path, array(
                                              'width' => 330,
                                              'height' => 300,
                                              'align' => 'right'
                                            ));

      // $this->addTableStyle('Main', $styleTable);
      // $table = $section->addTable('Main');
      //
      // // Row 3
      // $table->addRow(100);
      // $table->addCell(12000,  array('gridSpan' => 2, 'valign' => 'center', 'bgColor' => 'F2F2F2'))
      //       ->addText(htmlspecialchars($data->content->subheading), $subheadingFontStyle);
      //
      // $section->addImage($data->mainImage->path, array(
      //                                               'width' => 330,
      //                                               'height' => 300,
      //                                               'positioning' => 'absolute',
      //                                               'marginTop' => -5
      //                                             ));
      //
      // $last_key = 4;
      //
      // foreach($data->list as $key => $value) {
      //   $table->addRow(100);
      //   $cell = $table->addCell(1000,  array('bgColor' => 'F2F2F2'));
      //   $cell->addImage($data->iconPath->{$key});
      //   $cell = $table->addCell(10000,  array('bgColor' => 'F2F2F2'));
      //   $cell->addText(htmlspecialchars($value), $listFontStyle);
      // }
      //
      //
      //
      //
      $table->addRow(100);
      $table->addCell(12000, $styleMainSection)
            ->addText(htmlspecialchars($data->content->contentTitle), $subheadingFontStyle);

      $table->addRow(100);
      $table->addCell(12000,  $styleMainSection)
            ->addText(htmlspecialchars($data->content->contentDescription), $descriptionFontStyle);


      $this->addTableStyle('Main', $styleTable);
      $table = $section->addTable('Main');

      // Last Row
      $table->addRow(800);
      $table->addCell(12000,  array('gridSpan' => 2, 'valign' => 'center'))
            ->addImage("./resources/slide/image1.png", array(
              'align' => 'left',
              'width' => '100',
              'height' => '50'
            ));



      return $this;


    }
    public function save($obj, $path)
    {
      $oWriterDOCX = IOFactory::createWriter($obj, 'Word2007');
      $oWriterDOCX->save($path);
    }
}
