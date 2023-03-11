package com.app.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.app.excels.LibroDeTrabajo;

public class Application {

	public static void main(String[] args) {
		LibroDeTrabajo actual = new LibroDeTrabajo();
		File nuevoArchivoExcel = new File("C:\\Users\\fskurnik\\Desktop\\Franco\\uni\\Creando excel desde java",
				"Primer reporte desde Java.xls");
		File archivoLectura = new File("C:\\Users\\fskurnik\\Desktop\\Franco\\uni\\Creando excel desde java",
				"Nómina.xlsx");

		FileInputStream inputStream;
		ArrayList<String> nombres;

		try {
			inputStream = new FileInputStream(archivoLectura);
			XSSFWorkbook excelNomina = new XSSFWorkbook(inputStream);
			Sheet hojaNomina = excelNomina.getSheetAt(0);

			nombres = new ArrayList<String>();
			ArrayList<String> apellidos = new ArrayList<String>();
			ArrayList<Integer> legajos = new ArrayList<Integer>();
			
			for (int i = 1; i <= hojaNomina.getLastRowNum(); i++) {
				for (int j = 0; j < 1; j++) {
					apellidos.add(hojaNomina.getRow(i).getCell(j + 1).getStringCellValue());
					nombres.add(hojaNomina.getRow(i).getCell(j + 2).getStringCellValue());
					legajos.add((int) hojaNomina.getRow(i).getCell(j).getNumericCellValue());
				
				}
			}
			String encabezados[] = { "Nombre", "Apellido", "Legajo" };
			
			Sheet primeraHoja = actual.getActual().createSheet("Reporte 1");
			Row fila = primeraHoja.createRow(0);
			
			for (int i = 0; i < encabezados.length; i++) {
				Cell celdaDeLaFila = fila.createCell(i);
				celdaDeLaFila.setCellValue(encabezados[i]);

			}
			
			for(int i = 0; i < nombres.size(); i++) {
				Row filaNueva = primeraHoja.createRow(i + 1);
				for(int j = 0; j < 1; j++ ) {
					Cell celdaNombre = filaNueva.createCell(j);
					Cell celdaApellido = filaNueva.createCell(j + 1);
					Cell celdaLegajo = filaNueva.createCell(j + 2);
					celdaNombre.setCellValue(nombres.get(i));
					celdaApellido.setCellValue(apellidos.get(i));
					celdaLegajo.setCellValue(legajos.get(i));
				}
			}
			
			excelNomina.close();
		} catch (FileNotFoundException e1) {
			System.out.println("No existe el archivo");
			e1.printStackTrace();
		} catch (IOException e) {
			System.out.println("No funciona el archivo nómina");
			e.printStackTrace();
		}
		

		FileOutputStream salidaDeDatos;

		try {
			salidaDeDatos = new FileOutputStream(nuevoArchivoExcel);
			actual.escribirDatos(salidaDeDatos);
			actual.getActual().close();
			salidaDeDatos.close();
		} catch (IOException e) {
			System.err.println("No funciona el archivo");
			e.printStackTrace();

		}

	}

}
