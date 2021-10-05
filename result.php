<?php
function read_excel($filepath){
    ini_set('error_reporting', 0);
    ini_set('display_errors', 0);
    ini_set('display_startup_errors', 0);
    require_once "PHPExcel/Classes/PHPExcel.php"; //подключаем наш фреймворк

    $ar=array(); // инициализируем массив
    $inputFileType = PHPExcel_IOFactory::identify($filepath); // узнаем тип файла, excel может хранить файлы в разных форматах, xls, xlsx и другие
    $objReader = PHPExcel_IOFactory::createReader($inputFileType); // создаем объект для чтения файла
    $objPHPExcel = $objReader->load($filepath); // загружаем данные файла в объект
    $ar = $objPHPExcel->getActiveSheet()->toArray(); // выгружаем данные из объекта в массив
    return $ar; //возвращаем массив
}

function array_multisort_value(){
    $args = func_get_args();
    $data = array_shift($args);
    foreach ($args as $n => $field) {
        if (is_string($field)) {
            $tmp = array();
            foreach ($data as $key => $row) {
                $tmp[$key] = $row[$field];
            }
            $args[$n] = $tmp;
        }
    }
    $args[] = &$data;
    call_user_func_array('array_multisort', $args);
    return array_pop($args);
}

if ($_FILES) if ($_FILES['document']) $data = read_excel($_FILES['document']['tmp_name']); unset($data[0]);
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-F3w7mX95PdgyTmZZMECAngseQB83DfGTowi0iMjiWaeVhAn4FJkqJByhZMI3AhiU" crossorigin="anonymous">

    <title>Document</title>
</head>
<body>
    <table class="table">
        <thead>
            <tr>
                <td>Марка</td>
                <td>Артикул</td>
                <td>Количество</td>
            </tr>
        </thead>
        <tbody>
            <?php
            $array = [];
            $new_data = [];
            foreach ($data as $value) {
                if (isset($array[$value[0]][$value[1]]) and $array[$value[0]][$value[1]]) $array[$value[0]][$value[1]] += $value[2];
                else $array[$value[0]][$value[1]] = $value[2];
            }
            foreach ($array as $key => $value) {
                foreach ($value as $art => $qty) {
                    $new_data[] = array(
                        'mark' => $key, 
                        'art' => $art, 
                        'qty' => $qty
                    );
                }
            }
            foreach (array_multisort_value($new_data, 'qty', SORT_DESC) as $row) {
                ?>
                <tr>
                    <td><?= $row['mark'] ?></td>
                    <td><?= $row['art'] ?></td>
                    <td><?= $row['qty'] ?></td>
                </tr>
                <?php
            }
            ?>
        </tbody>
    </table>
</body>
</html>