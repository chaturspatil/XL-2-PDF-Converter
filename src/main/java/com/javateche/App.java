package com.javateche;

import com.itextpdf.text.FontFactory;
import com.javateche.xl2pdf.Excel2PdfPOI;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main1( String[] args )
    {
        System.out.println( "Hello World!" );

    }

    public static void main(String[] args) throws Exception{

        //String sExcelPath = "D:\\My_WORKSSSSSS\\TW-DAE-E2021_Excel-PDF\\evalTESTPWDS.xls";
        //String sPdfPath = "D:\\My_WORKSSSSSS\\TW-DAE-E2021_Excel-PDF\\evalTESTPWDS_pdf_java.pdf";

        //String sExcelPath = "D:\\My_WORKSSSSSS\\TW-DAE-E2021_Excel-PDF\\ReportType1 - Copy.xls";
        //String sPdfPath = "D:\\My_WORKSSSSSS\\TW-DAE-E2021_Excel-PDF\\ReportType1 - Copy_pdf.pdf";

        //String sExcelPath = "D:\\My_WORKSSSSSS\\TW-DAE-E2021_Excel-PDF\\NPCIL (1)\\NPCIL\\TENDER 2\\NPCIL_408960_FINBIDTEMP.xls";
        //String sPdfPath = "D:\\My_WORKSSSSSS\\TW-DAE-E2021_Excel-PDF\\NPCIL (1)\\NPCIL\\TENDER 2\\NPCIL_408960_FINBIDTEMP-pdf_10000-5000.pdf";

        //String sExcelPath = "D:\\My_WORKSSSSSS\\TW-DAE-E3-2022_Excel-PDF\\NPCIL TENDER 01\\NPCIL TENDER 01\\Individual Vendor Files\\Vendor4\\FINANCIALBID1680_202205301539020733.xls";
        //String sPdfPath = "D:\\My_WORKSSSSSS\\TW-DAE-E3-2022_Excel-PDF\\NPCIL TENDER 01\\NPCIL TENDER 01\\Individual Vendor Files\\Vendor4\\FINANCIALBID1680_202205301539020733.pdf";

        //String sExcelPath = "D:\\My_WORKSSSSSS\\TW-DAE-E3-2022_Excel-PDF\\NPCIL TENDER 01\\NPCIL TENDER 01\\Evaluation Sheets Generated after Opening\\Financial\\evalNPCIL_410004_FINANCIATEMPLATEBID1680.xls";
        //String sPdfPath = "D:\\My_WORKSSSSSS\\TW-DAE-E3-2022_Excel-PDF\\NPCIL TENDER 01\\NPCIL TENDER 01\\Evaluation Sheets Generated after Opening\\Financial\\evalNPCIL_410004_FINANCIATEMPLATEBID1680_pdf.pdf";

        String sExcelPath = "D:\\IntelJID\\WS\\XL-2-PDF-Converter\\samplefiles\\BOQADNDM.xls";
        String sPdfPath = "D:\\IntelJID\\WS\\XL-2-PDF-Converter\\samplefiles\\BOQADNDM_xls.pdf";


        //loadPdfFonts();
        //FontFactory.registerDirectory("D:\\My_WORKSSSSSS\\TW-DAE-E2021_Excel-PDF\\fonts", true);

        Excel2PdfPOI.convert2PDF(sExcelPath, sPdfPath);

        System.out.println("--------------done---------------------------");
    }

}
