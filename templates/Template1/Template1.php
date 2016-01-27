<?php
namespace Templates;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\IOFactory;

class Template1 extends PhpPresentation
{
    /*
     * As soon as the instance created, first slide also created.
     * So we need to call getActiveSlide() instead of createSlide().
     *
     * @return slide object
     */
    public function createFirstSlide()
    {
        return $this->getActiveSlide();
    }

    /*
     * Create new slide
     *
     * @return slide object
     */
    public function createNextSlide()
    {
      return $this->createSlide();
    }

    /*
     * Create Rich Text
     *
     * @return rich text
     */
    public function createRichText()
    {
      
    }



}
