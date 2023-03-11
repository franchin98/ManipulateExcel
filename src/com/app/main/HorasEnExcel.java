package com.app.main;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.LinkedList;
import java.util.List;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HorasEnExcel {

	public static void main(String[] args) {
		
				
		String rutaDelArchivo = JOptionPane.showInputDialog("Ingrese la ruta del archivo:");
		String nombreDelArchivo = JOptionPane.showInputDialog("Ingrese el nombre del archivo: ");

		BufferedInputStream bufferedInput = createBufferedWithFileInputStream(rutaDelArchivo, nombreDelArchivo);
		FileOutputStream bufferedOutput = createFileOutputStream();

		try {
			XSSFWorkbook archivoExcelDeLectura = new XSSFWorkbook(bufferedInput);
			Sheet hojaConDatos = archivoExcelDeLectura.getSheetAt(0);

			String tiempoPactado = "00:20:00";
			int conteo = 0;

			List<String> agentes, tiemposDeBaño;
			agentes = new LinkedList<String>();
			tiemposDeBaño = new LinkedList<String>();

			int conteoAgentes = 0;

			for (int i = 0; i < hojaConDatos.getLastRowNum() - 2; i++) {
				Row filaActual = hojaConDatos.getRow(i + 3);
				if (filaActual.getCell(30).getStringCellValue().compareTo(tiempoPactado) > 0
						&& filaActual.getRowNum() != hojaConDatos.getLastRowNum()) {
					System.out.println(
							(filaActual.getRowNum() + 1) + " - Agente: [" + filaActual.getCell(1).getStringCellValue()
									+ "] Tiempo en baño: " + filaActual.getCell(30).getStringCellValue());
					agentes.add(filaActual.getCell(1).getStringCellValue());
					tiemposDeBaño.add(filaActual.getCell(30).getStringCellValue());
					conteoAgentes++;
				}
				conteo++;
			}

			HSSFWorkbook reporte = new HSSFWorkbook();

			Sheet nuevaHoja = reporte
					.createSheet(LocalDate.now().getDayOfMonth() - 1 + "-0" + LocalDate.now().getMonthValue());

			String[] encabezados = { "Nombre agente", "Fecha", "Tiempo en baño" };

			Row primeraFila = nuevaHoja.createRow(0);
			for (int i = 0; i < encabezados.length; i++) {
				Cell celda = primeraFila.createCell(i);
				celda.setCellValue(encabezados[i]);
			}

			for (int i = 0; i < agentes.size(); i++) {
				Row filaActual = nuevaHoja.createRow(i + 1);
				Cell celdaNombreAgente = filaActual.createCell(0);
				Cell celdaFecha = filaActual.createCell(1);
				Cell celdaTiempoDeBaño = filaActual.createCell(2);

				celdaNombreAgente.setCellValue(agentes.get(i));
				celdaFecha.setCellValue(hojaConDatos.getRow(3).getCell(3).getNumericCellValue());
				celdaTiempoDeBaño.setCellValue(tiemposDeBaño.get(i));

			}

			System.out.println("Total de agentes encontrados: " + conteoAgentes);
			System.out.println("Total de filas: " + (conteo - 2));

			for (int i = 0; i < 3; i++)
				reporte.getSheetAt(0).autoSizeColumn(i);

			reporte.write(bufferedOutput);
			reporte.close();
			bufferedOutput.close();
			bufferedInput.close();
			archivoExcelDeLectura.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private static File nuevoArchivoExcel = new File(
			"C:\\Users\\fskurnik\\Desktop\\Franco\\uni\\Creando excel desde java", "Excedentes baño.xls");

	private static BufferedInputStream createBufferedWithFileInputStream(String ruta, String nombreArchivo) {
		try {

			File archivoLectura = new File(ruta, nombreArchivo);
			return new BufferedInputStream(new FileInputStream(archivoLectura));
		} catch (FileNotFoundException e) {
			System.err.println("Falla en método Create Buffered With File InputStream");
			e.printStackTrace();
		}

		return null;
	}

	private static FileOutputStream createFileOutputStream() {
		try {
			return new FileOutputStream(nuevoArchivoExcel);
		} catch (FileNotFoundException e) {
			System.err.println("Falla en método Create Buffered With File OutputStream");
			e.printStackTrace();
		}

		return null;
	}

}
