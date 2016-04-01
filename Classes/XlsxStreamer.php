<?php

namespace Snazzware;

class XlsxStreamer
{
    protected $dictionary = array();            // This will store every unique value we see during streaming.
    protected $dictionaryIndex = 0;             // Current length of the dictionary
    protected $rowIndex = 0;                    // Number of rows we have streamed out
    protected $zip = null;                      // ZipArchive instance
    protected $filePath = '';                   // Path to our ZipArchive output file
    protected $sheet = null;                    // File pointer for our worksheet
    protected $classDirectory = __DIR__;        // Path to the location of this class, used to reference template files.
    protected $sheetTemplate = null;            // Contents of our worksheet template file (loaded during open()

    /**
     * Creates a new zip file (with .xlsx extension) at the location specified by $filepath.
     *
     * @param string $filePath Filename to write output to. May include path, otherwise current working directory is used
     *
     * @return bool Returns true on success
     */
    public function open($filePath)
    {
        $result = true;

        // Store file path for later use
        $this->filePath = $filePath;

        // Reset variables
        $this->dictionary = array();
        $this->dictionaryIndex = 0;
        $this->rowIndex = 1;

        // Remove existing file if present
        if (file_exists($filePath)) {
            unlink($filePath);
        }

        // Create ZipArchive instance
        $this->zip = new \ZipArchive();
        if (!$this->zip->open($filePath, \ZIPARCHIVE::CREATE)) {
            $result = false;
        }

        if ($result) {
            // Create directories
            $this->zip->addEmptyDir('docProps');
            $this->zip->addEmptyDir('_rels');
            $this->zip->addEmptyDir('xl');
            $this->zip->addEmptyDir('xl/_rels');
            $this->zip->addEmptyDir('xl/theme');
            $this->zip->addEmptyDir('xl/worksheets');

            // List of template files that we will be using
            // Could have enumerated template directory contents, but I'm listing them out here for illustrative purposes
            $templateFiles = array(
                '[Content_Types].xml',
                'docProps/app.xml',
                'docProps/core.xml',
                'docProps/thumbnail.jpeg',
                '_rels/.rels',
                'xl/_rels/workbook.xml.rels',
                'xl/styles.xml',
                'xl/workbook.xml',
                'xl/theme/theme1.xml',
            );

            // Add each template file to the zip archive
            foreach ($templateFiles as $templateFile) {
                $this->zip->addFile("{$this->classDirectory}/template/{$templateFile}", $templateFile);
            }

            // Open up a temporary file where we will write out the sheet data
            $this->sheet = fopen($this->filePath.'_tmp_sheet1.xml', 'w');

            if ($this->sheet == false) {
                $result = false;
            } else {
                // Open the sheet template
                if ($this->sheetTemplate == null) {
                    $this->sheetTemplate = file_get_contents("{$this->classDirectory}/template/xl/worksheets/sheet1.xml");
                }

                // Determine the position of the <sheetData> open tag
                $sheetDataPos = strpos($this->sheetTemplate, '<sheetData>');

                // Copy everything up to and including the <sheetData> open tag from the template in to our sheet file
                fwrite($this->sheet, substr($this->sheetTemplate, 0, $sheetDataPos + strlen('<sheetData>')));
            }
        }

        return $result;
    }

    /**
     * Writes a row of data to the currently open xlsx file.
     *
     * @param array $row Array of non-objects (strings, ints, etc) to write to the file
     */
    public function write($row)
    {
        $col = 1;

        // Begin row
        fwrite($this->sheet, "<row r=\"{$this->rowIndex}\" spans=\"1:".count($row).'">');

        // Process each value
        foreach ($row as $value) {
            // If value is not already in our dictionary, add it and increment our dictionary index
            if (!isset($this->dictionary[$value])) {
                $index = $this->dictionaryIndex;
                $this->dictionary[$value] = $index;
                ++$this->dictionaryIndex;
            } else {
                // Obtain index of value from our dictionary
                $index = $this->dictionary[$value];
            }

            $colName = $this->eitoa($col); // convert column number in to excel column name

            // Output column to sheet file
            fwrite($this->sheet, "<c r=\"{$colName}{$this->rowIndex}\" t=\"s\"><v>{$index}</v></c>");

            // Increement our column counter
            ++$col;
        }

        // End row
        fwrite($this->sheet, '</row>');

        // Increment our row counter
        ++$this->rowIndex;
    }

    public function close()
    {
        // Obtain position of </sheetData> closing tag in our sheet template
        $sheetDataPos = strpos($this->sheetTemplate, '</sheetData>');

        // Copy everything from (and including) the </sheetData> tag from our sheet template to our sheet file
        fwrite($this->sheet, substr($this->sheetTemplate, $sheetDataPos));

        // Close our sheet file
        fclose($this->sheet);

        // Open temporary file to write our dictionary
        $f = fopen($this->filePath.'_tmp_sharedStrings.xml', 'w');

        // Write out header
        fwrite($f, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');

        // Write out sst open tag
        fwrite($f, "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{$this->dictionaryIndex}\" uniqueCount=\"{$this->dictionaryIndex}\">");

        // For each dictionary entry, write out entry
        foreach (array_keys($this->dictionary) as $dictionaryEntry) {
            fwrite($f, "<si><t>{$dictionaryEntry}</t></si>");
        }

        // Write closing sst tag
        fwrite($f, '</sst>');

        // Close our dictionary file
        fclose($f);

        // Add dictionary and sheet to our zip file
        $this->zip->addFile($this->filePath.'_tmp_sharedStrings.xml', 'xl/sharedStrings.xml');
        $this->zip->addFile($this->filePath.'_tmp_sheet1.xml', 'xl/worksheets/sheet1.xml');

        // Close the zip file
        $result = $this->zip->close();

        // Remove temporary files
        unlink($this->filePath.'_tmp_sharedStrings.xml');
        unlink($this->filePath.'_tmp_sheet1.xml');

        return $result;
    }

    /**
     * Converts a numeric value in to an excel column name (1 = A, 2 = B ... 26 = Z, 27 = AA, 28 = AB, etc...).
     *
     * @param int $num column number
     *
     * @return string excel column name
     */
    public function eitoa($num)
    {
        $numeric = ($num - 1) % 26;
        $letter = chr(65 + $numeric);
        $num2 = intval(($num - 1) / 26);
        if ($num2 > 0) {
            return $this->eitoa($num2).$letter;
        } else {
            return $letter;
        }
    }
}
