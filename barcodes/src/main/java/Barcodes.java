import net.sourceforge.barbecue.Barcode;
import net.sourceforge.barbecue.BarcodeException;
import net.sourceforge.barbecue.BarcodeFactory;
import net.sourceforge.barbecue.BarcodeImageHandler;
import net.sourceforge.barbecue.output.OutputException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;

public class Barcodes{

    private static final String FILE_NAME = "file.xlsx";

    public static void main(String[] args) {
        try {
            int row1 = 1;
            int row2 = 2;
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME)); //file
            Workbook workbook = new XSSFWorkbook(excelFile); //Make the file a workbook
            DataFormatter fmt = new DataFormatter(); //Get an instance of the formatter
            Sheet sheet = workbook.getSheetAt(0); //Get a sheet of the work book

            for(int i = 1; i <= 474; i++){
                Cell cell = workbook.getSheetAt(0).getRow(i).getCell(5); //cell 5 = NDC , increase row to go down the list
                String value = fmt.formatCellValue(cell);

                Barcode barcode = BarcodeFactory.createCode128(String.valueOf(value));
                File f =new File("barcode.png");
                BarcodeImageHandler.savePNG(barcode, f);

                InputStream is = new FileInputStream("barcode.png");
                byte[] bytes = IOUtils.toByteArray(is);
                int picID = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                is.close();

                CreationHelper help = workbook.getCreationHelper();

                Drawing draw = sheet.createDrawingPatriarch();

                ClientAnchor anchor = help.createClientAnchor();

                anchor.setCol1(9);
                anchor.setRow1(row1);
                anchor.setCol2(10);
                anchor.setRow2(row2);

                Picture pic = draw.createPicture(anchor, picID);
                row1++;
                row2++;
            }

            FileOutputStream fileOut = null;
            fileOut = new FileOutputStream("report.xlsx");
            workbook.write(fileOut);
            fileOut.close();


            System.out.println("Done");


        } catch (IOException | BarcodeException | OutputException e) {
            e.printStackTrace();
        }
    }
}