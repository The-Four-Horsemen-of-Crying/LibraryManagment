/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package testexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Enrique
 */
public class ExcelManager {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        ArrayList<StringBuffer> a;

        StringBuffer ab = new StringBuffer();

        ExcelManager e = new ExcelManager();

        String fileName = "test.xlsx";
        String path = "L:\\";
        String sheet = "test";

        a = e.readFile(fileName, path);
        e.show(a);

        System.out.println("\n\n*****************************************************\n\n");
        ab.append("13,14,15");
        a.add(ab);
        e.createFile(fileName, path, sheet, a);
        a = e.readFile(fileName, path);

        e.show(a);

    }
    
    public ArrayList<StringBuffer> readFile(String nombreArchivo, String rutaArchivo) {
        rutaArchivo += nombreArchivo;
        ArrayList<StringBuffer> document = new ArrayList<>();
        StringBuffer s = new StringBuffer();
        try (FileInputStream file = new FileInputStream(new File(rutaArchivo))) {
            // leer archivo excel
            XSSFWorkbook worbook = new XSSFWorkbook(file);
            //obtener la hoja que se va leer
            XSSFSheet sheet = worbook.getSheetAt(0);
            //obtener todas las filas de la hoja excel
            Iterator<Row> rowIterator = sheet.iterator();
            Row row;
            // se recorre cada fila hasta el final
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                //se obtiene las celdas por fila
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell;
                //se recorre cada celda
                while (cellIterator.hasNext()) {
                    // se obtiene la celda en específico y se la imprime
                    cell = cellIterator.next();
                    try {
                        s.append(cell.getStringCellValue());
                    } catch (Exception e) {
                        s.append(cell.getNumericCellValue());
                    }
                    if (cellIterator.hasNext()) {
                        s.append(",");
                    }
                }
                if (s.length() != 0) {
                    document.add(s);
                }
                s = new StringBuffer();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return document;
    }

    public void createFile(String nombreArchivo, String rutaArchivo, String hoja, ArrayList<StringBuffer> document) {
        rutaArchivo += nombreArchivo;
        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja1 = libro.createSheet(hoja);
        CellStyle style = libro.createCellStyle();
        Font font = libro.createFont();
        font.setBold(true);
        style.setFont(font);
        //generar los datos para el documento
        for (int i = 0; i < document.size(); i++) {
            XSSFRow row = hoja1.createRow(i);//se crea las filas
            String doc[] = document.get(i).toString().split(",");
            for (int j = 0; j < doc.length; j++) {
                if (i == 0) {//para la cabecera
                    XSSFCell cell = row.createCell(j);//se crea las celdas para la cabecera, junto con la posición
                    cell.setCellStyle(style); // se añade el style crea anteriormente
                    cell.setCellValue(doc[j]);//se añade el contenido
                } else {//para el contenido
                    XSSFCell cell = row.createCell(j);//se crea las celdas para la contenido, junto con la posición
                    cell.setCellValue(doc[j]); //se añade el contenido
                }
            }
        }

        File file = new File(rutaArchivo);
        try (FileOutputStream fileOuS = new FileOutputStream(file)) {
            if (file.exists()) {// si el archivo existe se elimina
                file.delete();
                System.out.println("Archivo eliminado");
            }
            libro.write(fileOuS);
            fileOuS.flush();
            fileOuS.close();
            System.out.println("Archivo Creado");

        } catch (FileNotFoundException e) {
            System.out.println("No se encontró el archivo o se encuentra abierto");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void show(ArrayList<StringBuffer> s) {
        for (StringBuffer stringBuffer : s) {
            System.out.println(stringBuffer.toString());
        }
    }

}
