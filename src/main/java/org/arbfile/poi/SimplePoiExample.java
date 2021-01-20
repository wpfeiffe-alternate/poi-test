package org.arbfile.poi;

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class SimplePoiExample
{
    public static void main(String[] args)
    {
        try
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet firstSheet = wb.createSheet();
            for (int i = 1; i <= 10; i++)
            {
                Row currRow = firstSheet.createRow(i - 1);
                for (int j = 1; j <= 10; j++)
                {
                    Cell currCell = currRow.createCell(j - 1);
                    currCell.setCellValue("Row " + i + ", Cell " + j);
                }
            }

            FileOutputStream out = new FileOutputStream("SimplePoiExample.xlsx");
            wb.write(out);
        }
        catch (Exception e)
        {
            System.out.println(e.getMessage());
        }
    }
}
