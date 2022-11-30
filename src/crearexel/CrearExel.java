/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package crearexel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class CrearExel {

    
    public static void main(String[] args) throws IOException {
        crearExel();
    }
    
    public  static  void  crearExel() throws IOException{
     Workbook book = new HSSFWorkbook();
        org.apache.poi.ss.usermodel.Sheet sheet =  book.createSheet("Biografia");
     
     Row row = sheet.createRow(01);
     row.createCell(0).setCellValue("Garcia Saldaña Gabriela Mariana");
     row.createCell(1).setCellValue(19);
     row.createCell(3).setCellValue(true);
     
     Cell celda = row.createCell(3);
     celda.setCellFormula(String.format("1+1",""));
     
     Row rouno = sheet.createRow(1);
     rouno.createCell(0).setCellValue(7);
     rouno.createCell(1).setCellValue(8);
     
     Cell celdados = rouno.createCell(2);
     celdados.setCellFormula(String.format("A%d+B%d", 2,2));
     
        try {
                FileOutputStream fileout = new FileOutputStream("Exel.xlsx");
                book.write(fileout);
                fileout.close();
         
        } catch (FileNotFoundException ex) {
            Logger.getLogger(CrearExel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }    
}
