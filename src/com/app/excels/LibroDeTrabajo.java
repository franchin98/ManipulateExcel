package com.app.excels;

import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;


public class LibroDeTrabajo {
	private HSSFWorkbook actual;
	
	
	public LibroDeTrabajo() {
		actual = new HSSFWorkbook();	
	}
	
	public HSSFWorkbook getActual() {
		return actual;
	}

	public void setActual(HSSFWorkbook actual) {
		this.actual = actual;
	}
	
	
	public void escribirDatos(OutputStream archivo) throws IOException {
		actual.write(archivo);
	}
}
