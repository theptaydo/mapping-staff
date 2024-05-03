/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mapping.staff;

import Staff.Misa;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author ADMIN
 */
public class MappingStaff {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            String excelFilePath = "C:\\Users\\ADMIN\\Downloads\\hosocanbonv.Xlsx";
            try (FileInputStream inputStream = new FileInputStream(new File(excelFilePath))) {
                Workbook workbook = WorkbookFactory.create(inputStream);
                Sheet sheet = workbook.getSheetAt(0); // Lấy sheet đầu tiên
                Iterator<Row> iterator = sheet.iterator();
                List<Misa> staffs = new ArrayList<>();
                Misa m = new Misa();
                while (iterator.hasNext()) {
                    Row nextRow = iterator.next();
                    Iterator<Cell> cellIterator = nextRow.cellIterator();
                    
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        int columnIndex = cell.getColumnIndex();
                        
                        switch (columnIndex) {
                            case 0:
                                m.setMaNV(cell.getStringCellValue());
                                break;
                            case 1:
                                m.setTen(cell.getStringCellValue());
                                break;
                            case 3:
                                m.setNgaySing(cell.getStringCellValue());
                                break;
                            case 4:
                                m.setNgayThuViec(cell.getStringCellValue());
                                break;
                            case 5:
                                m.setNgayNghiViec(cell.getStringCellValue());
                                break;
                            case 6:
                                m.setSdtLH(cell.getStringCellValue());
                                break;
                            case 7:
                                m.setEmailCT(cell.getStringCellValue());
                                break;
                            case 8:
                                break;
                            case 9:
                                m.setDiaChi(cell.getStringCellValue());
                                break;
                            case 10:
                                m.setViTriCV(cell.getStringCellValue());
                                break;
                            case 11:
                                m.setQLTT(cell.getStringCellValue());
                                break;
                        }
                    }
                    
                    // Tạo đối tượng Student và thêm vào danh sách
                    staffs.add(m);
                }   // In ra danh sách sinh viên
                for (Misa s : staffs) {
                    System.out.println(s);
                }
//            workbook.close();
            }
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }
    
}
