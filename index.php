
<button onclick="exportTableToCSV('cable_missing-03-06-2021.csv')" style="color:coral">Export To CSV File</button>
<br/><br/>

<?php 
/** Include path **/
set_include_path(get_include_path() . PATH_SEPARATOR . 'Classes/');
include 'PHPExcel/IOFactory.php';
$file = 'cable_missing-03-06-2021.xls';
$inputFileType = PHPExcel_IOFactory::identify($file);
$objReader = PHPExcel_IOFactory::createReader($inputFileType);
$objReader->setReadDataOnly(true);
$objPHPExcel = $objReader->load($file);
$objWorksheet = $objPHPExcel->getActiveSheet();
$CurrentWorkSheetIndex = 0;

foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
    // echo 'WorkSheet' . $CurrentWorkSheetIndex++ . "\n";
 echo '<strong>Worksheet number - </strong>', $objPHPExcel->getIndex($worksheet), PHP_EOL;
$lastRow = $worksheet->getHighestRow();
$colomncount = $worksheet->getHighestDataColumn();
$colomncount_number=PHPExcel_Cell::columnIndexFromString($colomncount);
    
echo "<table border='1'>";
	for($row=5;$row<=$lastRow-1;$row++){
		echo "<tr>";
		for($col=1;$col<=$colomncount_number-1;$col++){
			echo "<td>";
			echo $worksheet->getCell(PHPExcel_Cell::stringFromColumnIndex($col).$row)->getValue();
			echo "</td>";
		}
		echo "</tr>";
	}	
echo "</table>";
        
    }



?>



<script> 
function downloadCSV(csv, filename) {
    var csvFile;
    var downloadLink;

    // CSV file
    csvFile = new Blob([csv], {type: "text/csv"});

    // Download link
    downloadLink = document.createElement("a");

    // File name
    downloadLink.download = filename;

    // Create a link to the file
    downloadLink.href = window.URL.createObjectURL(csvFile);

    // Hide download link
    downloadLink.style.display = "none";

    // Add the link to DOM
    document.body.appendChild(downloadLink);

    // Click download link
    downloadLink.click();
}

function exportTableToCSV(filename) {
    var csv = [];
    var rows = document.querySelectorAll("table tr");
    
    for (var i = 0; i < rows.length; i++) {
        var row = [], cols = rows[i].querySelectorAll("td, th");
        
        for (var j = 0; j < cols.length; j++) 
            row.push(cols[j].innerText);
        
        csv.push(row.join(","));        
    }

    // Download CSV file
    downloadCSV(csv.join("\n"), filename);
}


</script>