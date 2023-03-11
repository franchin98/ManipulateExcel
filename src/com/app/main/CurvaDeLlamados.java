package com.app.main;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Hashtable;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CurvaDeLlamados {

	private static File archivoExportado = new File(
			"C:\\Users\\fskurnik\\Desktop\\Franco\\uni\\Creando excel desde java\\Curva de llamados",
			"Agentes por skill diario.xls");

	public static void main(String[] args) {

		String rutaArchivoDeMitrol = JOptionPane.showInputDialog(null, "Ingrese la ruta del archivo de mitrol:",
				"Curva de llamados", JOptionPane.INFORMATION_MESSAGE);

		String nombreArchivoDeMitrol = JOptionPane.showInputDialog(null, "Ingrese el nombre del archivo de mitrol:",
				"Curva de llamados", JOptionPane.INFORMATION_MESSAGE);
		
		IOUtils.setByteArrayMaxOverride(2000000000);
		BufferedInputStream archivoMitrol = createFileInputStream(rutaArchivoDeMitrol, nombreArchivoDeMitrol);
		FileOutputStream nuevoArchivoExcel = createFileOutputStream();

		try {
			
			XSSFWorkbook excelAux = new XSSFWorkbook(archivoMitrol);
			Sheet hoja = excelAux.getSheetAt(0);	

			Map<Integer, List<String>> agentes = new Hashtable<Integer, List<String>>();
			
			/*
			 * Este mapa nos va a permitir saber en cuántos intervalos aparece mediante
			 * el id de login del agente.  
			 *
			 */
			
			//Map<Integer, Linked> intervalosAgentes = new Hashtable<Integer, String>();
			
			
			for (int i = 0; i < hoja.getLastRowNum() - 2; i++) {
				Row row = hoja.getRow(i + 2);
				Integer key = Integer.parseInt(row.getCell(1).getStringCellValue());

				if (agentes.containsKey(key) && contieneElValor(row.getCell(3).getStringCellValue(), agentes.get(key))) 
					continue;
				else if(agentes.containsKey(key)){
					agentes.get(key).add(row.getCell(3).getStringCellValue());
				}

				int j = row.getRowNum();

				Row rowAux = hoja.getRow(j);
				LinkedList<String> campaniasAgente = new LinkedList<String>();
				
				while (key.equals(Integer.parseInt(rowAux.getCell(1).getStringCellValue()))) {
					campaniasAgente.add(rowAux.getCell(3).getStringCellValue());
					j++;
					rowAux = hoja.getRow(j);
				}

				// guardamos en un Mapa al LoginID como clave
				//
				agentes.put(key, campaniasAgente);

			}

			LinkedList<Integer> idAgentes = new LinkedList<Integer>();

			Set<Integer> IdAgentesAux = agentes.keySet();
			idAgentes.addAll(IdAgentesAux);

			HSSFWorkbook reporte = new HSSFWorkbook();
			Sheet hojaReporte = reporte.createSheet();

			Row encabezadoHojaReporte = hojaReporte.createRow(0);
			
			String campanias[] = { "Banco", "Banco Inversiones", "BIP Prestamos",
					"Cuenta DNI", "Cuenta DNI App Cobros", "Cuenta DNI Comercios", 
					"Cuenta DNI Datos CBU - 4 Digitos", "Cuenta DNI Programa Acompañar", 
					"E-Provincia Bapro", "Fraudes BAPRO", "Hipotecario Sin Validar",
					"Hipotecario Validado", "Mesa Ayuda Banca Internet", 
					"Mesa ayuda BIP empresas", "Opcion Premios", "Paquetes ENTRANTE CLIENTES", 
					"Productos y Servicios", "Reclamos", "Tarjeta Alimentar"} ;
			
			Cell encabezadoId = encabezadoHojaReporte.createCell(0);
			encabezadoId.setCellValue("Id Agente");
			
			for(int i = 1; i <= campanias.length; i++ ){
				Cell encabezadoCampania = encabezadoHojaReporte.createCell(i); 
				encabezadoCampania.setCellValue(campanias[i-1]);
			}

			for (Integer idActual : idAgentes) {
				Row fila = hojaReporte.createRow(hojaReporte.getLastRowNum() + 1);
				Cell celdaId = fila.createCell(0);
				celdaId.setCellValue(idActual.toString());

				LinkedList<String> campaniasAgente = (LinkedList<String>) agentes.get(idActual);

				for (String campaniaActual : campaniasAgente) {
					
					int indiceCelda = completarCeldaEnColumnaCorrespondiente(campaniaActual);
					
					Cell celdaCampania = fila.createCell(indiceCelda);
					
					celdaCampania.setCellValue("1");
				}

			}

			for (int i = 0; i < 19; i++)
				hojaReporte.autoSizeColumn(i);

			JOptionPane.showInternalMessageDialog(null, "Archivo exportado con éxito.",
					"Curva de llamados - by Skurnik Franco", JOptionPane.INFORMATION_MESSAGE);

			reporte.write(nuevoArchivoExcel);
			reporte.close();
			excelAux.close();
			archivoMitrol.close();
			nuevoArchivoExcel.close();

		} catch (IOException e) {

			e.printStackTrace();
		}

	}

	private static Integer completarCeldaEnColumnaCorrespondiente(String campaniaActual) {
		
		
		switch(campaniaActual) {
			case "Banco":
				return 1;
			case "Banco Inversiones":
				return 2;
			case "BIP Prestamos":
				return 3;
			case "Cuenta DNI":
				return 4;
			case "Cuenta DNI App Cobros":
				return 5;
			case "Cuenta DNI Comercios":
				return 6;
			case "Cuenta DNI Datos CBU - 4 Digitos":
				return 7;
			case "Cuenta DNI Programa Acompañar" :
				return 8;
			case "E-Provincia Bapro" :
				return 9;
			case "Fraudes BAPRO":
				return 10;
			case "Hipotecario Sin Validar":
				return 11;
			case "Hipotecario Validado":
				return 12;
			case "Mesa Ayuda Banca Internet":
				return 13;
			case "Mesa ayuda BIP empresas":
				return 14;
			case "Opcion Premios":
				return 15;
			case "Paquetes ENTRANTE CLIENTES": 
				return 16;
			case "Productos y Servicios":
				return 17;
			case "Reclamos":
				return 18;
			case "Tarjeta Alimentar":
				return 19;
			}
			
		return 19;
	}

	private static Boolean contieneElValor(String valor, List<String> lista) {
			
		return lista.contains(valor);
	}

	private static FileOutputStream createFileOutputStream() {
		try {
			return new FileOutputStream(archivoExportado);
		} catch (FileNotFoundException e) {
			System.err.println("Falla en método constructor de File OUTPUT");
			e.printStackTrace();
		}

		return null;
	}

	private static BufferedInputStream createFileInputStream(String rutaArchivoDeMitrol, String nombreArchivoDeMitrol) {
		try {
			return new BufferedInputStream(new FileInputStream(new File(rutaArchivoDeMitrol, nombreArchivoDeMitrol)));
		} catch (FileNotFoundException e) {
			System.err.println("Falla en metodo constructor de FileInputStream");
			e.printStackTrace();
		}
		return null;
	}

}
