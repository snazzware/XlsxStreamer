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
