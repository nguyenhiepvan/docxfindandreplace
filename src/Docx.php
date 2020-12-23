<?php

namespace Nguyenhiep\DocxFindAndReplace;

class Docx
{
    /**
     * @var $items array
     */
    private $items = [];

    /**
     * @var $templateFile string
     */
    protected $templateFile = '';

    /**
     * Docx constructor.
     * @param string $templateFile
     * @throws \Exception
     */
    public function __construct($templateFile)
    {
        if (!extension_loaded('zip')) {
            throw new \Exception('You must install the ZIP archive PHP extension.');
        }

        $this->templateFile = $templateFile;
    }

    /**
     * Set our docx template.
     *
     * @param string $templateFile
     * @return static
     * @throws \Exception
     */
    public static function create($templateFile)
    {
        if (!file_exists($templateFile)) {
            throw new \Exception('Template file not found.');
        }

        return new static($templateFile);
    }

    /**
     * Wrap our items to replace with curly braces.
     *
     * @param array $items
     * @return $this
     * @throws \Exception
     */
    public function replace($items = [])
    {
        if (!is_array($items)) {
            throw new \Exception('Items must be listed in an array.');
        }

        $wrappedItems = [];
        foreach ($items as $key => $item) {
            $wrappedItems[$key] = $item;
        }
        $this->items = $wrappedItems;

        return $this;
    }

    /**
     * Create & save our new docx file from the template.
     *
     * @param string $destinationFile
     * @throws \Exception
     */
    public function save($destinationFile)
    {
        if (empty($this->templateFile)) {
            throw new \Exception('No template file set.');
        }

        copy($this->templateFile, $this->templateFile . '.tmp');

        $this->createFromZip(new \ZipArchive());

        copy($this->templateFile . '.tmp', $destinationFile);

        unlink($this->templateFile . '.tmp');
    }

    /**
     * Grab the contents of the template, find and replace, and create a new temp docx file.
     *
     * @param \ZipArchive $zip
     * @return void
     */
    private function createFromZip(\ZipArchive $zip)
    {
        $zip->open($this->templateFile . '.tmp');

        $contents = $zip->getFromName($name = "word/document.xml");
        $contents = $this->replacing($contents);
        $zip->deleteName($name);
        $zip->addFromString($name, $contents);

        $contents = $zip->getFromName($name = 'word/_rels/document.xml.rels');
        $contents = $this->replacing($contents);
        $zip->deleteName($name);
        $zip->addFromString($name, $contents);
        $zip->close();
    }

    private function replacing($contents){
        foreach ($this->items as $key => $value) {
            if ($this->isRegularExpression($key)) {
                $contents = preg_replace($key, $value, $contents);
            } else {
                $contents = str_replace($key, $value, $contents);

            }
        }
        return
            $contents;
    }

    private function isRegularExpression($string)
    {
        return @preg_match($string, '') !== FALSE;
    }
}
