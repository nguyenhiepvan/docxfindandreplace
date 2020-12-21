<?php

use \Nguyenhiep\DocxFindAndReplace\Docx;

class DocxTest extends PHPUnit_Framework_TestCase
{
    public function test_docx_is_created()
    {
        Docx::create(__DIR__ . "/test.docx")->replace(
            [
                'FIRSTNAME' => 'Joe',
                'LASTNAME' => 'Nguyenhiep',
                'ADDRESS1' => '123 Main St',
                'CITY' => 'Maple Heights',
                'STATE' => 'OH',
                'ZIPCODE' => '44137'
            ]
        )->save(__DIR__ . '/joeNguyenhiep.docx');

        $this->assertFileExists(__DIR__ . '/joeNguyenhiep.docx');
    }

    public function test_file_found_and_replaced()
    {
        $contents =  $this->getDocxContents(__DIR__ . '/joeNguyenhiep.docx');

        $this->assertContains('Joe Nguyenhiep', $contents);
        $this->assertContains('123 Main St', $contents);
        $this->assertContains('Maple Heights', $contents);
        $this->assertContains('OH', $contents);
        $this->assertContains('44137', $contents);
    }

    public function test_temp_file_does_not_exist()
    {
        $this->assertFalse(file_exists(__DIR__ . '/test.docx.tmp'));
    }

    public function test_file_is_deleted()
    {
        unlink(__DIR__ . '/joeNguyenhiep.docx');

        $this->assertFalse(file_exists(__DIR__ . '/joeNguyenhiep.docx'));
    }

    private function getDocxContents($file)
    {
        $zip = new \ZipArchive();
        $zip->open($file);

        return $zip->getFromName("word/document.xml");
    }
}