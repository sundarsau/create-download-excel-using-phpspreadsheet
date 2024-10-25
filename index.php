<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

include 'cfg/dbconnect.php';
$sql = "SELECT * FROM users order by name";
$stmt = $conn->prepare($sql);
$stmt->execute();
$result = $stmt->get_result();

if ($result->num_rows> 0){
	$spreadsheet = new Spreadsheet(); // create an object of Spreadsheet
	$activeWorksheet = $spreadsheet->getActiveSheet(); // active worksheet

	$activeWorksheet->setCellValue('A1', 'ID');
	$activeWorksheet->setCellValue('B1', 'Name');
	$activeWorksheet->setCellValue('C1', 'Email');
	$activeWorksheet->setCellValue('D1', 'Age');
	$activeWorksheet->setCellValue('E1', 'Gender');

	$row = 1; $col = 0;

	while($value = $result->fetch_assoc()){
		$row++;
		$col= 65;  // column A
		$activeWorksheet->setCellValue(chr($col).$row, $value['id']);
		$col++; // column B
		$activeWorksheet->setCellValue(chr($col) . $row, $value['name']);
		$col++; // column C
		$activeWorksheet->setCellValue(chr($col) . $row, $value['email']);
		$col++; // column D
		$activeWorksheet->setCellValue(chr($col) . $row, $value['age']);
		$col++; // column E
		$activeWorksheet->setCellValue(chr($col) . $row, $value['gender']);
	}
	$writer = new Xlsx($spreadsheet);
	$filename = "users_".time().".xlsx";
	$writer->save("output/". $filename);

}
?>
<!DOCTYPE html>
<html lang="en">

<head>
	<title>Create & download Excel using PhpSpreadsheet</title>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
	<link rel="stylesheet" href="css/style.css">
</head>

<body>
	<div class="container">
		<h1>List of Users</h1>
		<div class="table-responsive">
			<table class="table table-bordered table-hover table-striped">
				<thead class="table-dark">
					<tr>
						<th>ID</th>
						<th>Name</th>
						<th>Email</th>
						<th>Age</th>
						<th>Gender</th>
					</tr>
				</thead>
				<tbody>
					<?php if ($result->num_rows > 0) {
						foreach ($result as $row) {  ?>
							<tr>
								<td>
									<?php echo $row['id']; ?>
								</td>
								<td>
									<?php echo $row['name']; ?>
								</td>
								<td>
									<?php echo $row['email']; ?>
								</td>
								<td>
									<?php echo $row['age']; ?>
								</td>
								<td>
									<?php echo $row['gender']; ?>
								</td>
							</tr>
						<?php
						}
					} else { ?>
						<tr>
							<td colspan="5"> No Users Found</td>
						</tr>
					<?php }	?>
				</tbody>
			</table>
		</div>
		<?php if ($result->num_rows > 0) {?>
			<div class="float-end mb-5">
				<a class="btn btn-success" href="output/<?php echo $filename;?>">Download Excel</a>
			</div>
		<?php } ?>
	</div>	
</body>

</html>