# XlsxStreamer

Simple class to "stream" data to a .xlsx file.

The .xlsx file is created on disk, except for the dictionary. This allows far less memory consumption than using a solution such as (the excellent) PHPExcel library.

I plan to add more capabilities to this streamer in the future, such as basic formatting, support for embedded hyperlinks, etc.

Basic Usage
-----------
    <?php

    require __DIR__.'/vendor/autoload.php';

    $streamer = new \Snazzware\XlsxStreamer();

    $streamer->open('test.xlsx');
    $streamer->write([
        'This is A1',
        'This is B1',
    ]);
    $streamer->write([
        'This is A2',
        'This is B2',
    ]);
    $streamer->close();


Blog posts
----------

http://blog.snazzware.com/2016/01/05/streaming-large-data-sets-to-excel-xlsx-targets-part-1/

http://blog.snazzware.com/2016/01/09/streaming-large-data-sets-to-excel-xlsx-targets-part-2/
