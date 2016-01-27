<?php
namespace Templates;

interface TemplateInterface
{
  public function create($json);
  public function save($obj, $path);
}
