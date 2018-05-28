<?php

	ini_set('max_execution_time', 2400); //300 seconds = 5 minutes
    ini_set('memory_limit', '3512M');
	require_once './PHPExcel/Classes/PHPExcel.php';
	$sexo = array('', 'Masculino', 'Femenino');
	$Alumnos = json_decode($_POST["arregloAlumnos"],true);
	$SQLEstandar = $_POST["SQL"];
    $informacionBD = json_decode($_POST["informacionBD"],true);

    if (!isset($Alumnos)){
        header('location: http://www.saludescolar.yucatan.gob.mx/');
    }
    
	$numeroAlumnos = count($Alumnos);
	$indexArregloAlumnos = 0;
	echo $numeroAlumnos;
	$numeroElementos = $numeroAlumnos;

	$objPHPExcel = new PHPExcel();
    $objPHPExcel->getProperties()->setCreator("Sicoy")
                                         ->setLastModifiedBy("Sicoy")
                                         ->setTitle("Alumnos Medicion")
                                         ->setSubject("Alumnos Medicion")
                                         ->setDescription("Documento Creado desde Aplicacion Sicoyv21")
                                         ->setKeywords("office 2007 openxml php")
                                         ->setCategory("Test result file");
	$mysqli = new mysqli($informacionBD["servidor"], $informacionBD["usuario"], $informacionBD["password"], $informacionBD["nombreBD"]);
    if(!$mysqli){
        echo "No se pudo realizar la conexión PHP - MySQL"; 
    } else {

		$cont = 1;
		$objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A1', 'Ciclo')
                    ->setCellValue('B1', 'Periodo')
                    ->setCellValue('C1', 'Alumno')
                    ->setCellValue('D1', 'Curp')
                    ->setCellValue('E1', 'Sexo')
                    ->setCellValue('F1', 'Edad Meses')
                    ->setCellValue('G1', 'Peso')
                    ->setCellValue('H1', 'Estatura')
                    ->setCellValue('I1', 'Cintura')
                    ->setCellValue('J1', 'Imc')
                    ->setCellValue('K1', 'Puntuacion Z')
                    ->setCellValue('L1', 'Percentil Imc')
                    ->setCellValue('M1', 'Percentil Cintura')
                    ->setCellValue('N1', 'Percentil Estatura')
                    ->setCellValue('O1', 'Interpretacion E.F.')
                    ->setCellValue('P1', 'Interpretacion Antiguo')
                    ->setCellValue('Q1', 'Fuerza Brazos')
                    ->setCellValue('R1', '# Abdominales')
                    ->setCellValue('S1', 'Flexion Tronco')
                    ->setCellValue('T1', 'Resistencia')
                    ->setCellValue('U1', 'Flexibilidad Pie Izq')
                    ->setCellValue('V1', 'Flexibilidad Pie Der')
                    ->setCellValue('W1', 'Actividades Sedentarias')
                    ->setCellValue('X1', 'Actividades Activas')
                    ->setCellValue('Y1', 'Estado Fisico Puntos')
                    ->setCellValue('Z1', 'Porcentaje Grasa')
                    ->setCellValue('AA1', 'Flexibilidad')
                    ->setCellValue('AB1', 'Course Navette')
                    ->setCellValue('AC1', 'Salto Horizontal')
                    ->setCellValue('AD1', 'Nivel')
                    ->setCellValue('AE1', 'Clave Institucion')
                    ->setCellValue('AF1', 'Institucion')
                    ->setCellValue('AG1', 'Grado')
                    ->setCellValue('AH1', 'Grupo')
                    ->setCellValue('AI1', 'Municipio');
        $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('S')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('T')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('U')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('X')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('AD')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('AF')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('AG')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('AH')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setAutoSize(true);
		

		for ($i=0; $i < $numeroElementos; $i++) { 
			$sql = $SQLEstandar." and m.id_alumno_medicion = '".$Alumnos[$indexArregloAlumnos]["id_alumno_medicion"]."'";
		
			$indexArregloAlumnos++;
			
    		$resultado = $mysqli->query($sql);
    		if ($resultado->num_rows > 0){
                $fila =  $resultado->fetch_array(MYSQLI_ASSOC);
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'.($cont+1), $fila["ciclo"])
                    ->setCellValue('B'.($cont+1), $fila["periodo"])
                    ->setCellValue('C'.($cont+1), utf8_encode($fila["nombre_completo"]))
                    ->setCellValue('D'.($cont+1), $fila["curp"])
                    ->setCellValue('E'.($cont+1), $sexo[$fila["sexo"]])
                    ->setCellValue('F'.($cont+1), $fila["edad_meses"])
                    ->setCellValue('G'.($cont+1), $fila["peso"])
                    ->setCellValue('H'.($cont+1), $fila["estatura"])
                    ->setCellValue('I'.($cont+1), $fila["circunferencia_cintura"])
                    ->setCellValue('J'.($cont+1), $fila["imc"])
                    ->setCellValue('K'.($cont+1), $fila["puntuacion_z"])
                    ->setCellValue('L'.($cont+1), $fila["percentil_imc"])
                    ->setCellValue('M'.($cont+1), $fila["percentil_cintura"])
                    ->setCellValue('N'.($cont+1), $fila["percentil_estatura"])
                    ->setCellValue('O'.($cont+1), utf8_encode($fila["interpretacion_z"]))
                    ->setCellValue('P'.($cont+1), utf8_encode($fila["interpretacion_imc"]))
                    ->setCellValue('Q'.($cont+1), $fila["fuerza_brazos"])
                    ->setCellValue('R'.($cont+1), $fila["abdominales"])
                    ->setCellValue('S'.($cont+1), $fila["flexion_tronco"])
                    ->setCellValue('T'.($cont+1), $fila["resistencia"])
                    ->setCellValue('U'.($cont+1), $fila["flexibilidad_pie_izq"])
                    ->setCellValue('V'.($cont+1), $fila["flexibilidad_pie_der"])
                    ->setCellValue('W'.($cont+1), $fila["actividades_sedentarias"])
                    ->setCellValue('X'.($cont+1), $fila["actividades_activas"])
                    ->setCellValue('Y'.($cont+1), $fila["estado_fisico_puntos"])
                    ->setCellValue('Z'.($cont+1), $fila["porcentaje_grasa"])
                    ->setCellValue('AA'.($cont+1), $fila["flexibilidad"])
                    ->setCellValue('AB'.($cont+1), $fila["course_navette"])
                    ->setCellValue('AC'.($cont+1), $fila["salto_horizontal"])
                    ->setCellValue('AD'.($cont+1), utf8_encode($fila["nivel"]))
                    ->setCellValue('AE'.($cont+1), utf8_encode($fila["clave_institucion"]))
                    ->setCellValue('AF'.($cont+1), utf8_encode($fila["institucion"]))
                    ->setCellValue('AG'.($cont+1), $fila["grado"])
                    ->setCellValue('AH'.($cont+1), $fila["grupo"])
                    ->setCellValue('AI'.($cont+1), utf8_encode($fila["municipio"])); 
                $cont++;
	        }
	        
		}

		$objPHPExcel->getActiveSheet()->setTitle('Alumnos Medicion');

		ob_clean();
		
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="ExportaConsultaAgrupado.xlsx"');
        header('Cache-Control: max-age=0');
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save('php://output');
   
    }

?>