
<?php

		function out($sql, $name2, $name3, $link){

			if($result = $link->query($sql)){
   				$rowsCount = $result->num_rows;
    			echo '<table cellspacing="2" border="1" cellpadding="5"><tr><th><p>Дата</p></th><th><p>' . $name2 . '</p></th><th><p>' . $name3 . '</p></th></tr>';
    			foreach($result as $row){
       				echo "<tr>";
           				echo "<td>" . $row["1"] . "</td>";
            			echo "<td>" . $row["2"] . "</td>";
            			echo "<td>" . $row["3"] . "</td>";
        			echo "</tr>";
    			}
    			echo "</table>";
    			$result->free();
				}
			}


		require_once 'Classes/PHPExcel.php';
		$link = mysqli_connect("localhost", "root", "password", "db");

	
		$tmpfname = "Задание.xlsx";
		$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
		$excelObj = $excelReader->load($tmpfname);
		$worksheet = $excelObj->getSheet(0);
		$highestRow = $worksheet->getHighestRow();
		$highestColumn = $worksheet->getHighestColumn();
		$highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn) - 3;

		echo '<table cellspacing="2" border="1" cellpadding="5">' . "\n";
		for ($row = 4; $row <= $highestRow; ++$row) {

			$id = $worksheet->getCell("A".$row)->getValue();
        	$com_name = $worksheet->getCell("B".$row)->getValue();
        	$fact_qliq_data1 = $worksheet->getCell("C".$row)->getValue();
        	$fact_qliq_data2 = $worksheet->getCell("D".$row)->getValue();
       	 	$fact_qoil_data1 = $worksheet->getCell("E".$row)->getValue();
       	 	$fact_qoil_data2 = $worksheet->getCell("F".$row)->getValue();
        	$forecast_qliq_data1 = $worksheet->getCell("G".$row)->getValue();
        	$forecast_qliq_data2= $worksheet->getCell("H".$row)->getValue();
       	 	$forecast_qoil_data1 = $worksheet->getCell("I".$row)->getValue();
       	 	$forecast_qoil_data2 = $worksheet->getCell("J".$row)->getValue();
        	$c_date = $worksheet->getCell("K".$row)->getValue();
        	$c_date = date('ymd', \PHPExcel_Shared_Date::ExcelToPHP($c_date));

        	$sql = mysqli_query($link, "Insert into test (id, com_name, fact_qliq_data1, fact_qliq_data2, fact_qoil_data1, fact_qoil_data2, forecast_qliq_data1, forecast_qliq_data2, forecast_qoil_data1,forecast_qoil_data2, c_date) value ($id,'$com_name', $fact_qliq_data1, $fact_qliq_data2, $fact_qoil_data1, $fact_qoil_data2, $forecast_qliq_data1, $forecast_qliq_data2, $forecast_qoil_data1, $forecast_qoil_data2, $c_date)");

    	
   			/*if ($sql) {
    			echo '<p>Данные успешно добавлены в таблицу.</p>';
   			 } else {
     			echo '<p>Произошла ошибка: ' . mysqli_error($link) . '</p>';
    			}*/
    		}


		if ($link == false){
		    print("Ошибка: Невозможно подключиться к MySQL " . mysqli_connect_error());
		}
		else {
		    print("\nСоединение установлено успешно!");
		}

		$sql = "select c_date as '1', SUM(fact_qliq_data1 + fact_qliq_data2) as '2', SUM(forecast_qliq_data1 + forecast_qliq_data2) as '3' from test
			where com_name = 'company1' group by c_date";
		out($sql, 'Фактические данные по qliq компании company1', 'Ожидаемые данные по qliq компании company1', $link);


		$sql = "select c_date as '1', SUM(fact_qoil_data1 + fact_qoil_data2) as '2', SUM(forecast_qoil_data1 + forecast_qoil_data2) as '3' from test where com_name = 'company1' group by c_date";
		out($sql, 'Фактические данные по qoil компании company1', 'Ожидаемые данные по qoil компании company1', $link);


		$sql = "select c_date as '1', SUM(fact_qliq_data1 + fact_qliq_data2) as '2', SUM(forecast_qliq_data1 + forecast_qliq_data2) as '3' from test
			where com_name = 'company2' group by c_date";
		out($sql, 'Фактические данные по qliq компании company2', 'Ожидаемые данные по qliq компании company2', $link);


		$sql = "select c_date as '1', SUM(fact_qoil_data1 + fact_qoil_data2) as '2', SUM(forecast_qoil_data1 + forecast_qoil_data2) as '3' from test where com_name = 'company2' group by c_date";
		out($sql, 'Фактические данные по qoil компании company2', 'Ожидаемые данные по qoil компании company2', $link);
?>




