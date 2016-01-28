<?php
namespace Templates\Word;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;



class Template1 extends PhpWord implements \Templates\TemplateInterface
{
    public function create($json)
    {
      $data = json_decode(utf8_encode($json));

      /* Note: any element you append to a document must reside inside of a Section. */

      // Adding an empty Section to the document...
      $section = $this->addSection();
      // Adding Text element to the Section having font styled by default...
      $section->addText(
          htmlspecialchars($data->content->title)
      );

      // $section->addPageBreak();

      $section->addText(
          htmlspecialchars($data->content->contentDescription)
      );

      return $this;


    }
    public function save($obj, $path)
    {
      $oWriterDOCX = IOFactory::createWriter($obj, 'Word2007');
      $oWriterDOCX->save($path);
    }
}
