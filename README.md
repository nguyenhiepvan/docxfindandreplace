# DOCX Find & Replace

This is a simple find and replace utility for DOCX files. Simple way to take a DOCX template, map some variables, and save a new copy.

## Installation

Via Composer

``` bash
$ composer require Nguyenhiep/docxfindandreplace
```

## Usage
In your DOCX template you will need to wrap any variables you would like to replace with curly braces (e.g. ``firstname``). You can use regex expressions as key
``` php
\Nguyenhiep\DocxFindAndReplace\Docx::create(__DIR__ . "/template.docx")->replace(
    [
        "firstname"                                         => "nguyen",
        "lastname"                                          => "hiep",
        "/[-0-9a-zA-Z.+_]+@[-0-9a-zA-Z.+_]+.[a-zA-Z]{2,4}/" => "nguyenhiepvan.bka@gmail.com"
    ]
)->save(__DIR__ . '/newfile.docx');
```
