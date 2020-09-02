/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package testexcel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
 
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Enrique
 */
public class TestExcel {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        String fileName = "test.xlsx";
        String path = "L:\\" + fileName;
        String sheet = "test";
        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja1 = libro.createSheet(sheet);
        String [] header= new String[]{"1", "2"};
        String[][] document = new String[][]{
            {"hola", "mundo"},
            {"version", "excel"},
            {"jodanse",""},
            {"Lora","Estuvo aquí"}
        };

        for (int i = 0; i < document.length; i++) {
            XSSFRow row = hoja1.createRow(i);//se crea las filas
            for (int j = 0; j < header.length; j++) {
                XSSFCell cell = row.createCell(j);//se crea las celdas para la contenido, junto con la posición
                cell.setCellValue(document[i][j]); //se añade el contenido

            }
        }
        
    File file;
		file = new File(path);
		try (FileOutputStream fileOuS = new FileOutputStream(file)){						
			if (file.exists()) {// si el archivo existe se elimina
				file.delete();
				System.out.println("Archivo eliminado");
			}
			libro.write(fileOuS);
			fileOuS.flush();
			fileOuS.close();
			System.out.println("Archivo Creado");
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}
		
 
	}

}
