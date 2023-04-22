<?php
#ini_set('display_errors','on'); version_compare(PHP_VERSION, '5.5.0') <= 0 ? error_reporting(E_WARNING & ~E_NOTICE & ~E_DEPRECATED) : error_reporting(E_ALL & ~E_NOTICE & ~E_DEPRECATED & ~E_STRICT);   # DEBUGGING
ini_set('display_errors','on'); error_reporting(E_ALL); # STRICT DEVELOPMENT



/**
 * [REQUERIMIENTOS]
 * git clone https://github.com/PHPOffice/PHPExcel.git
 * path libreria/
 */



# Verificar si se ha enviado el formulario
if (isset($_POST['submit'])) {
    ini_set('max_execution_time', 0);
set_time_limit(0);

echo '<pre>';
    # Verificar si se ha subido un archivo
    if (isset($_FILES['archivo']) && $_FILES['archivo']['error'] === UPLOAD_ERR_OK) {
        
        # Obtener la extensión del archivo
        $ext = pathinfo($_FILES['archivo']['name'], PATHINFO_EXTENSION);
        
        # Verificar si es un archivo Excel
        if ($ext === 'xlsx' || $ext === 'xls') {
            
            # Incluir la librería PHPExcel
            require_once '../libreria/PHPExcel/Classes/PHPExcel.php';
            require_once '../../config.inc.php';
            define("_DB_", $dbconfig['db_name']);
            define("_USER_", $dbconfig['db_username']);
            define("_PASS_", $dbconfig['db_password']);
            define("_HOST_", $dbconfig['db_server']);
            
            #print_r($dbconfig);
            #die();

            # Crear un objeto PHPExcel para leer el archivo Excel
            $objPHPExcel = PHPExcel_IOFactory::load($_FILES['archivo']['tmp_name']);

            # Obtener la hoja activa del archivo Excel
            $worksheet = $objPHPExcel->getActiveSheet();

            # Obtener el número de filas y columnas de la hoja
            $numFilas = $worksheet->getHighestRow();
            echo('[INFO]: Cantidad de filas '.$numFilas).'<br>';

            $numCols = $worksheet->getHighestColumn();
            echo('[INFO]: Ultima Columnas '.$numCols).'<br>';

            try {
                $numColsIndex = PHPExcel_Cell::columnIndexFromString($numCols);
            } catch (Exception $e) {
                echo "Error al obtener el índice de la última columna: " . $e->getMessage();
                exit;
            }
            echo('[INFO]: Cantidad de Columnas '.$numColsIndex).'<br>';


            # Crear la definición de la columna para la primary key
            $campo = 'ID INT(11) NOT NULL AUTO_INCREMENT PRIMARY KEY';

            # Crear la definición de las demás columnas
            for ($i = 1; $i <= $numColsIndex; $i++) {
              $letra = PHPExcel_Cell::stringFromColumnIndex($i);
              $campo .= ", temp$letra TEXT NULL";
              $campoInsert .= " temp$letra ,";
            }
            $campoInsert = substr($campoInsert, 0, -1);

            # Ejecutar para crear la tabla si existe la tabla la borrara y la volvera a crear
            $nombreTabla = 'a_import_usuario_' . date('Y_m');

            crearTabla($nombreTabla, [$campo]);

            
            # Iterar por todas las filas del archivo Excel
            for ($row = 1; $row <= (int)$numFilas; $row++) {

                # Obtener los valores de las celdas de la fila
                $rowData = $worksheet->rangeToArray('A' . $row . ':' . $numCols . $row, null, true, false);

                # Generar el insert
                $query.= "INSERT INTO $nombreTabla ($campoInsert) ";
                # Insertar cada fila en la tabla
                foreach ($rowData as $arrayz) {
                    $fields_str = "'" . implode("', '", $arrayz) . "'";
                    $query.= " VALUES ($fields_str); ";

                    # Validar que la fila tenga el número correcto de columnas
                    if (count($row) != $numColsIndex) {
                        continue;
                    }


                }

            }

            # Iniciar la conexión a la base de datos utilizando PDO
            $pdo = new PDO("mysql:host="._HOST_.";dbname="._DB_.";charset=utf8", _USER_, _PASS_,);
            $stmt = $pdo->prepare($query);
            $stmt->execute();
            
            # Cerrar la conexión a la base de datos
            $pdo = null;
            
            # Mostrar un mensaje de éxito
            echo('[INFO]: Los datos han sido guardados en la base de datos '.$numColsIndex).'<br>';
            
        } else {
            echo('[ERROR]: El archivo debe ser un Excel (.xlsx o .xls)').'<br>';
        }
        
    } else {
            echo('[ERROR]: Debe seleccionar un archivo para subir').'<br>';
            echo "";
    }

    echo '</pre>';
}


#------------FUNCIONES REUTILIZABLES


function eliminarTabla($nombreTabla) {
    try {
        # Iniciar la conexión a la base de datos utilizando PDO
        $opciones = [
            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_EMULATE_PREPARES => false,
        ];
        try {
            $pdo = new PDO("mysql:host="._HOST_.";dbname="._DB_.";charset=utf8", _USER_, _PASS_, $opciones);
        } catch (PDOException $e) {
            die("Error al conectar: " . $e->getMessage());
        }

        # Preparar la sentencia SQL para eliminar la tabla
        $sql = "DROP TABLE IF EXISTS $nombreTabla";
        echo('[INFO]: '.$sql).'<br>';

        # Ejecutar la sentencia SQL para eliminar la tabla
        $pdo->exec($sql);

        # Cerrar la conexión a la base de datos
        $pdo = null;

        return true;
    } catch(PDOException $e) {
        echo "Error al eliminar la tabla: " . $e->getMessage();
        return false;
    }
}


function crearTabla($nombreTabla, $campos) {
    try {
        # Eliminar la tabla si existe
        eliminarTabla($nombreTabla);

        # Iniciar la conexión a la base de datos utilizando PDO
        $opciones = [
            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_EMULATE_PREPARES => false,
        ];
        try {
            $pdo = new PDO("mysql:host="._HOST_.";dbname="._DB_.";charset=utf8", _USER_, _PASS_, $opciones);
        } catch (PDOException $e) {
            die("Error al conectar: " . $e->getMessage());
        }

        # Preparar la sentencia SQL para crear la tabla
        $sql = "CREATE TABLE $nombreTabla (";
        foreach ($campos as $campo) {
            $sql .= "$campo, ";
        }
        $sql = rtrim($sql, ", ") . ")";

        # Ejecutar la sentencia SQL para crear la tabla
        $pdo->exec($sql);

        # Cerrar la conexión a la base de datos
        $pdo = null;
        echo('[INFO]: TABLA CREADA '.$nombreTabla).'<br>';

        return true;
    } catch(PDOException $e) {
        echo "Error al crear la tabla: " . $e->getMessage();
        return false;
    }
}

?>


<!DOCTYPE html>
<html>
<head>
    <title>Subir archivo Excel</title>
</head>
<body>
    <h1>Subir archivo Excel</h1>
    <form action="" method="post" enctype="multipart/form-data">
        <input type="file" name="archivo" required>
        <br>
        <input type="submit" name="submit" value="Subir archivo">
    </form>
</body>
</html>
