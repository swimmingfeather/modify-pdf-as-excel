package org.example;

//import org.apache.pdfbox.cos.COSName;
//import org.apache.pdfbox.cos.COSString;
//import org.apache.pdfbox.pdfparser.PDFStreamParser;
//import org.apache.pdfbox.pdfwriter.ContentStreamWriter;
import org.apache.pdfbox.pdmodel.PDDocument;
//import org.apache.pdfbox.pdmodel.PDPage;
//import org.apache.pdfbox.pdmodel.common.PDStream;
import org.apache.pdfbox.pdmodel.interactive.form.PDField;
import org.apache.pdfbox.printing.PDFPageable;
import org.apache.poi.ss.usermodel.*;

import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * http://pdfbox.apache.org/
 *
 * @author fish
 *
 */
public class PDF_Edit {
    private String pdf_dir;
    private String excel_dir;
    private String save_dir;

    private InputStream ins;
    private Workbook wb;
    private Sheet sheet;

    private PDDocument helloDocument;
//    private PDPage firstPage;
//    private PDStream updatedStream;
//    private OutputStream out;
//    private ContentStreamWriter tokenWriter;
//    private PDFStreamParser parser;
//    private List tokens;

    public PDF_Edit(String pdf, String excel){
        this.pdf_dir = pdf;
        this.excel_dir = excel;

        readExcel();
        editPDF();

    }
    public  void readExcel(){

        try {
            ins=new FileInputStream(new File(this.excel_dir));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        try {
            wb = WorkbookFactory.create(ins);
        } catch (Exception e) {
            e.printStackTrace();
        }
        try {
            ins.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        this.sheet = wb.getSheetAt(0);
    }

//    public void replaceString(int i){
//
//        for(Object o:tokens) {
//            if(o instanceof COSString) {
//                COSString cs = (COSString) o;
//                String string = cs.getString();
//
//                if(-1 == string.indexOf("column"))
//                    continue;
//
//                DataFormatter dataFormatter = new DataFormatter();
//                for(int k = 1; k <= sheet.getRow((short)0).getPhysicalNumberOfCells();k++) {
//                    System.out.println(string);
//                    string = string.replaceFirst("column"+k, dataFormatter.formatCellValue(sheet.getRow(i).getCell((short) k)));
//                }
//
//                cs.setValue(string.getBytes());
//            }
//        }
//
//
////        for (int j = 0; j < tokens.size(); j++)
////        {
////            Object next = tokens.get(j);
////            if (next instanceof Operator)
////            {
////                Operator op = (Operator) next;
////                // Tj and TJ are the two operators that display strings in a PDF
////                if (op.getName().equals("Tj"))
////                {
////                    // Tj takes one operator and that is the string
////                    // to display so lets update that operator
////                    COSString previous = (COSString) tokens.get(j - 1);
////                    String string = previous.getString();
////
////                    if(-1 == string.indexOf("column"))
////                        continue;
////
////
////                    string = previous.getString();
////                    DataFormatter dataFormatter = new DataFormatter();
////                    for(int k = 1; k <= this.sheet.getRow((short)0).getPhysicalNumberOfCells();k++) {
////                        System.out.println(dataFormatter.formatCellValue(sheet.getRow(i).getCell((short) k)));
////                        string = string.replaceFirst("column"+k, dataFormatter.formatCellValue(sheet.getRow(i).getCell((short) k)));
////                    }
////                    previous.setValue(string.getBytes());
////
////
////                }
////
////            }
////        }
//    }


    public void editPDF() {

        try {
            PrinterJob job = PrinterJob.getPrinterJob();
            job.printDialog();

            // pdfwithText
            for(int i = 1;i <= this.sheet.getLastRowNum(); i++) {
                helloDocument = PDDocument.load(new File(this.pdf_dir));
                //firstPage = helloDocument.getPage(0);
                //updatedStream = new PDStream(helloDocument);
//                parser = new PDFStreamParser(firstPage);
//                parser.parse();
//                tokens = parser.getTokens();
//                out = updatedStream.createOutputStream(COSName.FLATE_DECODE);
//                tokenWriter = new ContentStreamWriter(out);

                Map<String, String> map = new HashMap<>();
                DataFormatter dataFormatter = new DataFormatter();
                for(int k = 1; k <= sheet.getRow((short)0).getPhysicalNumberOfCells();k++) {
                    map.put("column"+k, dataFormatter.formatCellValue(sheet.getRow(i).getCell((short) k)));
                }

                List<PDField> fields = helloDocument.getDocumentCatalog().getAcroForm().getFields();
                for (PDField field : fields) {
                    for (Map.Entry<String, String> entry : map.entrySet()) {
                        if (entry.getKey().equals(field.getFullyQualifiedName())) {
                            field.setValue(entry.getValue());
                            field.setReadOnly(true);
                        }
                    }
                }


                job.setPageable(new PDFPageable(helloDocument));
                job.setJobName(i +".pdf");
                job.print();
//                File out = new File("out.pdf");
//                helloDocument.save(out);
//                helloDocument.close();
                //this.replaceString(i);
//                tokenWriter.writeTokens(tokens);
//                firstPage.setContents(updatedStream);
               // out.close();
                //helloDocument.save(this.save_dir + '/' + i +".pdf"); //Output file name
                helloDocument.close();
            }
            // now that the tokens are updated we will replace the page content stream.

        } catch (IOException | PrinterException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }



}