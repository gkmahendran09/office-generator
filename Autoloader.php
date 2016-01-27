<?php
ini_set('error_reporting', E_ERROR); // E_ALL, E_NOTICE, E_WARNING, E_ERROR
ini_set('display_errors', On);  // On or Off

// Load Common files
require_once __DIR__ . '/libs/PhpOffice/Common/Autoloader.php';
\PhpOffice\Common\Autoloader::register();

// Load files for PowerPoint
require_once __DIR__ . '/libs/PhpOffice/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();

// Load files for Word
require_once __DIR__ . '/libs/PhpOffice/PhpWord/Autoloader.php';
\PhpOffice\PhpWord\Autoloader::register();
