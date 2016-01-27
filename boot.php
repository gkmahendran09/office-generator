<?php
ini_set('error_reporting', E_ERROR); // E_ALL, E_NOTICE, E_WARNING, E_ERROR
ini_set('display_errors', On);  // On or Off

// Autoload library files
require_once __DIR__ . '/libs/Common/Autoloader.php';
\PhpOffice\Common\Autoloader::register();

require_once __DIR__ . '/libs/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();
