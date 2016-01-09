<?php

require_once("XlsxStreamer.php");

// Create a random data set
$rows = array();
for ($i=0;$i<100;$i++) {
	$row = array();
	for ($j=0;$j<50;$j++) {
		$row[] = rand();
	}
	$rows[] = $row;
}

// Start the streamer
$streamer = new XlsxStreamer();
if ($streamer->open("test.xlsx")) {

	// Write out the data set
	foreach ($rows as $row) {
		$streamer->write($row);
	}

	// Close the streamer
	$streamer->close();
} else {
	echo "Unable to create output file.";
}

?>