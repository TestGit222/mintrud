<?

/* Для сайта - добавить ограничение доступа, разные папки и т.п. */

if ($_POST['upl']==1) {

    $uploaddir = 'file/';
    $uploadfile = $uploaddir . 'user.xlsx'; // Это для локального сайта, для интернет - измените имя
    if (move_uploaded_file($_FILES['userfile']['tmp_name'], $uploadfile)) {
    //echo "Ваш файл успешно загрузился на сервер!";
    $is_upload=true;
    } else {
        echo "[ОШИБКА: Файл не загрузился на сервер!]";
    }
}
else if ($_GET['del']=='del'){
    
    unlink('file/user.xlsx');
    unlink('file/user.xml');
    
}
    
function f_isPassed($str) {
    
    if($str=='удовлетворительно') $out='true';
    else $out='false';
    
    return $out;
    
}

function f_date($str) {
    
    $dateArr=explode("/",$str);
    
    $out=$dateArr[2].'-'.$dateArr[1].'-'.$dateArr[0];
    
    return $out;
    
}

?><!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Генератор xml для Минтруда</title>
<link href="excel.css" rel="stylesheet" type="text/css">
</head>

<body>
<div class="main">
<div class="cont">
<h1>Генератор xml для Минтруда</h1>
<?
if (is_file("file/user.xlsx")) {
    
require_once 'excel/Classes/PHPExcel.php';
$excel = PHPExcel_IOFactory::load("file/user.xlsx");

$xls=$excel ->getSheet(0)->toArray();
    
foreach($xls as $row){
    $y++;
    $x=0;
    foreach($row as $col){
        $x++;
        $val[$y][$x]=$col;   
    }
}

print('
<table class="listUser">
  <tbody>
    <tr>
      <td>Фамилия</td>
      <td>Имя</td>
      <td>Отчество</td>
      <td>СНИЛС</td>
      <td>Должность</td>
      <td>ИНН организации</td>
      <td>Организация</td>
      <td>ИНН образовательного центра</td>
      <td>Полное название образовательного центра</td>
      <td>Программа (номер)</td>
      <td>Программа (название)</td>
      <td>Результат тестирования (true/false)</td>
        <td>Дата (год-месяц-день)</td>
        <td>Номер протокола</td>
    </tr>');

      
for($yy=2;$yy<=$y;$yy++) {
    
    $fio=explode(' ',$val[$yy][1]);
    $LastName[$yy]=$fio[0];
    $FirstName[$yy]=$fio[1];
    $MiddleName[$yy]=$fio[2];
    $Snils[$yy]=$val[$yy][2];
    $Position[$yy]=$val[$yy][3];
    $EmployerInn[$yy]=$val[$yy][4];
    $EmployerTitle[$yy]=$val[$yy][5];
    $Inn[$yy]=$val[$yy][6];
    $Title[$yy]=$val[$yy][7];
    $isPassed[$yy]=f_isPassed($val[$yy][9]);
    $program=explode(' ',$val[$yy][8],2);
    $learnProgramId[$yy]=$program[0];
    $LearnProgramTitle[$yy]=$program[1];
    $Date[$yy]=f_date($val[$yy][10]);
    $ProtocolNumber[$yy]=$val[$yy][11];
        
    print('<tr><td>'.$LastName[$yy].'</td><td>'.$FirstName[$yy].'</td><td>'.$MiddleName[$yy].'</td><td>'.$Snils[$yy].'</td><td>'.$Position[$yy].'</td><td>'.$EmployerInn[$yy].'</td><td>'.$EmployerTitle[$yy].'</td><td>'.$Inn[$yy].'</td><td>'.$Title[$yy].'</td><td>'.$learnProgramId[$yy].'</td><td>'.$LearnProgramTitle[$yy].'</td><td>'.$isPassed[$yy].'</td><td>'.$Date[$yy].'</td><td>'.$ProtocolNumber[$yy].'</td></tr>');
    
}  
      
if ($is_upload==true){ // Была загрузка - формируем XML
    
    $xml='<?xml version="1.0" encoding="utf-8"?>
<RegistrySet>';
    for($yy=2;$yy<=$y;$yy++) {
        $xml=$xml.'
    <RegistryRecord>
        <Worker>
            <LastName>'.$LastName[$yy].'</LastName>
			<FirstName>'.$FirstName[$yy].'</FirstName>
			<MiddleName>'.$MiddleName[$yy].'</MiddleName>
			<Snils>'.$Snils[$yy].'</Snils>
			<Position>'.$Position[$yy].'</Position>
			<EmployerInn>'.$EmployerInn[$yy].'</EmployerInn>
			<EmployerTitle>'.$EmployerTitle[$yy].'</EmployerTitle>
        </Worker>
		<Organization>
			  <Inn>'.$Inn[$yy].'</Inn>
              <Title>'.$Title[$yy].'</Title>
        </Organization>
		<Test isPassed="true" learnProgramId="'.$learnProgramId[$yy].'">
		      <Date>2024-09-24</Date>
              <ProtocolNumber>'.$ProtocolNumber[$yy].'</ProtocolNumber>
              <LearnProgramTitle>'.$LearnProgramTitle[$yy].'</LearnProgramTitle>
        </Test>
    </RegistryRecord>';
        
    }
    
    $xml=$xml.'
</RegistrySet>';
    
    $handle = fopen($uploaddir."user.xml", "w");
    fwrite($handle,$xml);
    fclose($handle);
    
}
      
print('</tbody></table>');
}
?>    
  
<h2 class="hr">Инструкции для пользователя</h2>
<blockquote>
      <p>1. Загрузите файл сюда (обрабатывается только первый лист):</p>
</blockquote>
<blockquote><form action="main.php" method="post" enctype="multipart/form-data" name="form1" id="form1">
      Загрузить:   
        <input type="file" name="userfile" id="userfile">
      <input type="submit" name="submit" id="submit" value="Отправить"><br>
      <strong>В InternetExplorer файлы не загружаются!</strong>
        <input type="hidden" name="upl" value="1">
</form></blockquote>
    <blockquote>
      <p>2. Проверьте правильность распознавания по таблице вверху.</p>
      <p>3. <a href="file/user.xml">Скачайте полученный XML</a>.</p>
      <p>4. <a href="main.php?del=del">Удалите загруженные файлы</a> (иначе они останутся в общем доступе).</p>
    </blockquote>
</div>
</div>
</body>
</html>
