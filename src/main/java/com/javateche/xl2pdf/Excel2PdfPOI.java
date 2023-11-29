package com.javateche.xl2pdf;

import java.io.*;
import java.lang.reflect.Field;
import java.time.format.DateTimeFormatter;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.hssf.util.HSSFColor;

import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.*;

public class Excel2PdfPOI {

	public static final String IS_COL_SPAN = "isColSpan";
	public static final String IS_ROW_SPAN = "isRowSpan";
	public static final String COL_START_IND = "colStartInd";
	public static final String COL_END_IND = "colEndInd";
	public static final String ROW_START_IND = "rowStartInd";
	public static final String ROW_END_IND = "rowEndInd";
	
	public static final int HEIGHT_DIVIDE_BY = 10;
	
	private static Map<String, Font> mFonts = new HashMap<String, Font>();
		

	public static void convert2PDF(String sExcelPath, String sPdfPath) throws Exception {
				
		FileInputStream input_document = new FileInputStream(new File(sExcelPath));
		
		HSSFWorkbook workbook = new HSSFWorkbook(input_document); 
		
		Map<Integer, Object> xlProps = getExcelProperties(workbook);
		
		if(xlProps.size() > 0) {
			Set<Float> sPageSige = (Set<Float>) xlProps.get(100000);
			System.out.println("---------------------------------------"+sPageSige);
		}
		
				
		int totalFonts = workbook.getNumberOfFontsAsInt();
		Set<String> sfonts = new HashSet<String>();
		
		for(int i = 0; i <= totalFonts; i++) {
			HSSFFont font = workbook.getFontAt(i);
		
			sfonts.add(font.getFontName());
		}
		System.out.println(sfonts);
		
		Iterator<Sheet> sheetIterator = workbook.iterator();
		
		Document pdfDoc = new Document(PageSize.A0);
	//	Document pdfDoc = new Document(new Rectangle(10000, 5000));
		//pdfDoc.setPageSize(PageSize.A1);

		PdfWriter.getInstance(pdfDoc, new FileOutputStream(sPdfPath));
		pdfDoc.open();

		//Loop through sheet
		while(sheetIterator.hasNext()) {
			pdfDoc.newPage();
			// Read worksheet into HSSFSheet
			HSSFSheet sheet = (HSSFSheet) sheetIterator.next();//my_xls_workbook.getSheetAt(0); 
			
			Map<String, Object> mCellProps = new HashMap<String, Object>();
			
			Map<Integer, Integer> mSpanDetails = new HashMap<Integer, Integer>();
			Set<String> sMergdCells = new HashSet<String>();
			
			if(workbook.isSheetVeryHidden(workbook.getSheetIndex(sheet.getSheetName()))) {

				int iMaxCol = getMaxColumns(sheet);
								
				if(iMaxCol > 0) {

					List<CellRangeAddress> lMergedCells = sheet.getMergedRegions();

					// To iterate over the rows
					Iterator<Row> rowIterator = sheet.iterator();
					
					float [] fArrColWidth = getColumnWidth(sheet, iMaxCol);
					//no of columnes
					PdfPTable pdfTable = new PdfPTable(fArrColWidth);
					pdfTable.setWidthPercentage(100);
					//Loop through rows.
					//int iCellToDeduct = 0;
					while(rowIterator.hasNext()) {
						//iCellToDeduct = 0;
						Row row = rowIterator.next(); 
						//Iterator<Cell> cellIterator = row.cellIterator();
						
						//We will use the object below to dynamically add new data to the table
						PdfPCell pdfCell = new PdfPCell(new Phrase(" "));;
		
						//while(cellIterator.hasNext()) {
						//for(int i = 0; i < iMaxCol; i++) {
						int i = 0;
						while(i < iMaxCol ) {//(iMaxCol - iCellToDeduct)
							Cell cell = row.getCell(i); //cellIterator.next(); 
							if(cell != null) {
								if(sMergdCells.contains(cell.getRowIndex()+","+cell.getColumnIndex())) {
									i++;
									continue;
								}
							
								mCellProps.clear();
								if(isCellMerged(lMergedCells, mCellProps, cell)) {
									
									if(mCellProps.get(IS_COL_SPAN) != null && (boolean) mCellProps.get(IS_COL_SPAN)) {	
										int iColSpaned = ((int)mCellProps.get(COL_END_IND)) - ((int)mCellProps.get(COL_START_IND)) + 1;
										//table_cell =  new PdfPCell();
										while(i <= (int)mCellProps.get(COL_END_IND)) {											
											cell = row.getCell(i);
											
											if(cell != null ) {//&& cell.getStringCellValue().length() > 0												
												//table_cell = new PdfPCell(new Phrase(cell.getStringCellValue()));// toString()));	
												pdfCell = new PdfPCell();
												
												setPdfCellProps(cell, pdfCell, workbook);
												
												//pdfCell = new PdfPCell(new Phrase(getCellValue(cell)));
												pdfCell.setColspan(((int)mCellProps.get(COL_END_IND)) - ((int)mCellProps.get(COL_START_IND)) + 1);																
																																															
												i = (int)mCellProps.get(COL_END_IND) + 1;
												break;												
											}else {											
												i++;
											}
										}
										if(mCellProps.get(IS_ROW_SPAN) != null && (boolean) mCellProps.get(IS_ROW_SPAN)) {
											int iRowSpaned = ((int)mCellProps.get(ROW_END_IND)) - ((int)mCellProps.get(ROW_START_IND)) + 1;
											
											pdfCell.setRowspan(((int)mCellProps.get(ROW_END_IND)) - ((int)mCellProps.get(ROW_START_IND)) + 1);	
											
											int iNextRowInd = cell.getRowIndex();
											int iCnt = 1;

											while(iCnt < iRowSpaned) { 
												iNextRowInd = iNextRowInd + 1;
												if(mSpanDetails.containsKey(iNextRowInd)) { 
													mSpanDetails.put(iNextRowInd, mSpanDetails.get(iNextRowInd) + 1 +(iColSpaned -1)); 
												}else { 
													mSpanDetails.put(iNextRowInd, 1 +(iColSpaned -1));
												}
												int iCurntColInd = cell.getColumnIndex();
												for(int j=0; j<iColSpaned; j++) {
													sMergdCells.add(iNextRowInd+","+iCurntColInd++);
												}
												
												iCnt++;
											}
											
										}
									}else if(mCellProps.get(IS_ROW_SPAN) != null && (boolean) mCellProps.get(IS_ROW_SPAN)){
										
										//cell = row.getCell(i);
										if(cell != null ) {//&& !sMergdCells.contains(cell.getRowIndex()+","+cell.getColumnIndex())
											int iRowSpaned = ((int)mCellProps.get(ROW_END_IND)) - ((int)mCellProps.get(ROW_START_IND)) + 1;
											pdfCell = new PdfPCell();

											setPdfCellProps(cell, pdfCell, workbook);

											pdfCell.setRowspan(iRowSpaned);

											int iNextRowInd = cell.getRowIndex();
											int iCnt = 1;

											while(iCnt < iRowSpaned) {
												
												iNextRowInd = iNextRowInd + 1;
												if(mSpanDetails.containsKey(iNextRowInd)) { 
													mSpanDetails.put(iNextRowInd, mSpanDetails.get(iNextRowInd) + 1); 
												}else { 
													mSpanDetails.put(iNextRowInd, 1);
												}
												sMergdCells.add(iNextRowInd+","+cell.getColumnIndex());
												iCnt++;
											}											
										}
										i++;
									}else {
										i++;
									}									
								} else {						
									//table_cell = new PdfPCell(new Phrase(cell.getStringCellValue()));// toString()));
									//pdfCell = new PdfPCell(new Phrase(getCellValue(cell)));
									
									pdfCell = new PdfPCell();
									setPdfCellProps(cell, pdfCell, workbook);
									
									//setCellValue(cell, table_cell);								
									//setBackgroundColor(cell, pdfCell);									
									i++;
								}
							
							}else {
								pdfCell=new PdfPCell(new Phrase(" "));
								pdfCell.setBorder(0);
								i++;
							}
							pdfTable.addCell(pdfCell);
						}
					}

					//Finally add the table to PDF document
					pdfDoc.add(pdfTable);      

				}
			}
		}

		pdfDoc.close();                
		//we created our pdf file..
		input_document.close(); //close xls
	}

	private static Map<Integer, Object> getExcelProperties(HSSFWorkbook workbook) throws Exception{
		Map<Integer, Object> excelProps = new HashMap<Integer, Object>();
		
		Set<Float> sPageSige = new HashSet<Float>();
		
		Iterator<Sheet> sheetIterator = workbook.iterator();
		int i = 0;
		while(sheetIterator.hasNext()) {
			HSSFSheet sheet = (HSSFSheet) sheetIterator.next();//my_xls_workbook.getSheetAt(0); 
			if(!workbook.isSheetVeryHidden(workbook.getSheetIndex(sheet.getSheetName()))) {
				Map<String, Object> sheetProps = new HashMap<String, Object>();
				
				int iMaxCol = getMaxColumns(sheet);
				float [] fArrColWidth = getColumnWidth(sheet, iMaxCol);
				
				float fTotalWidth = 0;
				
				for(float fColWidth : fArrColWidth) {
					fTotalWidth += (fColWidth * 200);
				}
				
				sheetProps.put("MaxCol", iMaxCol);
				sheetProps.put("ColWidthArr", fArrColWidth);
				
				sPageSige.add(fTotalWidth);
				
				excelProps.put(++i, sheetProps);
			}
		}
		excelProps.put(100000, sPageSige);
		
		return excelProps;
	}
	
	private static int getMaxColumns(HSSFSheet sheet) {
		FileInputStream input_document;
		Set<Integer> sColCount = new HashSet<Integer>();
		Iterator<Row> rowIterator = sheet.iterator();

		while(rowIterator.hasNext()) {
			Row row = rowIterator.next(); 
			sColCount.add(row.getPhysicalNumberOfCells());
		}
		
		if(sColCount.size() > 0) {
			return Collections.max(sColCount);
		}else {
			return 0;
		}

	}
	
	private static float [] getColumnWidth(HSSFSheet sheet, int iColCnt){
		float [] fArrColWidth = new float[iColCnt];
		
		for(int i = 0; i < iColCnt; i++) {
			fArrColWidth[i] = sheet.getColumnWidth(i)/200 ;
		}
		return fArrColWidth;
	}

	private static boolean isCellMerged(List<CellRangeAddress> lMergedCells, Map<String, Object> mCellProps, Cell cell) {
		mCellProps.clear();
		if(lMergedCells.size() > 0) {
			for(CellRangeAddress cellRngArrd : lMergedCells) {
				if(cellRngArrd.isInRange(cell)) {
					if(cellRngArrd.getFirstColumn() < cellRngArrd.getLastColumn()) {
						mCellProps.put(IS_COL_SPAN, true);
						mCellProps.put(COL_START_IND, cellRngArrd.getFirstColumn());
						mCellProps.put(COL_END_IND, cellRngArrd.getLastColumn());
					}

					if(cellRngArrd.getFirstRow() < cellRngArrd.getLastRow()) {
						mCellProps.put(IS_ROW_SPAN, true);
						mCellProps.put(ROW_START_IND, cellRngArrd.getFirstRow());
						mCellProps.put(ROW_END_IND, cellRngArrd.getLastRow());
					}

					return true;
					//break;	
				}
			}
		}
		return false;
	}
	
	private static void setPdfCellProps(Cell cell, PdfPCell pdfCell, HSSFWorkbook workBook) throws IOException, Exception {
		String sCellVal = getCellValue(cell);
			
		//Phrase p = setFont(cell, pdfCell, sCellVal, workBook);
		//pdfCell.setPhrase(p);
	
		Phrase p = setRichFont(cell, pdfCell, sCellVal, workBook);
		pdfCell.setPhrase(p);
		
		setBackgroundColor(cell, pdfCell);
		setAlignment(cell, pdfCell);
		setCellBorder(cell, pdfCell);
	}
	
	private static void setBackgroundColor(Cell cell, PdfPCell pdfCell) {
		CellStyle cellStyle = cell.getCellStyle();
	//	cellStyle.getFontIndexAsInt();
		short [] rgb = ((HSSFColor)cellStyle.getFillForegroundColorColor()).getTriplet();
		if(rgb[0] == 0 && rgb[1] == 0 && rgb[2] == 0) {
			rgb[0] = rgb[1] = rgb[2] =  255;
		}
		pdfCell.setBackgroundColor(new BaseColor(rgb[0], rgb[1], rgb[2]));
	}
	
	private static String getCellValue(Cell cell) {
		String sCellVal = "";
		
		switch(cell.getCellType()) {
			case NUMERIC:
				if(DateUtil.isCellDateFormatted(cell)){
				//if (HSSFDateUtil.isCellDateFormatted(cell)) {
					DateTimeFormatter formatter_1 = DateTimeFormatter.ofPattern("dd-MM-yyyy HH:mm");
					sCellVal = cell.getLocalDateTimeCellValue().format(formatter_1);
				//	sCellVal = cell.getLocalDateTimeCellValue().toString();
				}else {

					cell.setCellType(CellType.STRING);
					//pdfCell = new PdfPCell(new Phrase(cell.getStringCellValue()));// toString()));
					sCellVal = cell.getStringCellValue();
				}
				break;
	
			case STRING:				
				//table_cell = new PdfPCell(new Phrase(cell.getStringCellValue()));
				sCellVal = cell.getStringCellValue();
				break;
	
			case BOOLEAN:				
				cell.setCellType(CellType.STRING);
				//table_cell = new PdfPCell(new Phrase(cell.getStringCellValue()));// toString()));	
				sCellVal = cell.getStringCellValue();
				break;
	
			default: 
				cell.setCellType(CellType.STRING);
				//table_cell = new PdfPCell(new Phrase(cell.getStringCellValue()));// toString()));	
				sCellVal = cell.getStringCellValue();
				break;							
		}
		return sCellVal;
	}
	
	private static Phrase setFont(Cell cell, PdfPCell pdfCell, String sCellVal, HSSFWorkbook workBook) throws Exception, IOException {
		//CellStyle cellStyle = cell.getCellStyle();
		//int iFontInd = cellStyle.getFontIndexAsInt() == 0 ? 30 : cellStyle.getFontIndexAsInt();
				
		int fontIndex = cell.getCellStyle().getFontIndexAsInt();
		
		HSSFFont excelFont = workBook.getFontAt(fontIndex);
		
		boolean isBold = excelFont.getBold();
		
		String sFontName = excelFont.getFontName();
				
		short sHeight = excelFont.getFontHeight();
		
		boolean isItalic = excelFont.getItalic();
		
		Font font = FontFactory.getFont(sFontName, sHeight/HEIGHT_DIVIDE_BY);
		//font.setSize(sHeight/20);
		
		if(isBold && isItalic) {
			font.setStyle(Font.BOLDITALIC);
		}else if(isBold){
			//font.setStyle(Font.BOLD);
			font.setStyle(Font.NORMAL);
		}else if(isItalic) {
			font.setStyle(Font.ITALIC);
		}else {
			font.setStyle(Font.NORMAL);
		}
		
		if(excelFont.getHSSFColor(workBook) != null) {
			
			short [] rgb = ((HSSFColor)excelFont.getHSSFColor(workBook)).getTriplet();
		
			if(rgb[0] == 255 && rgb[1] == 255 && rgb[2] == 255) {
				rgb[0] = rgb[1] = rgb[2] = 0;
			}
			
			font.setColor(rgb[0], rgb[1], rgb[2]);			
		}
		return new Phrase(System.lineSeparator() + sCellVal + System.lineSeparator()+System.lineSeparator(), font);
	}
	
	public static Phrase setRichFont(Cell cell, PdfPCell pdfCell, String sCellVal, HSSFWorkbook workBook) {
		Phrase ph = new Phrase();
		
		switch(cell.getCellType()) {

		case STRING:
			RichTextString richtextstring = cell.getRichStringCellValue();

			String textstring = richtextstring.getString();

			if (richtextstring.numFormattingRuns() > 0) { // we have formatted runs

				//Phrase ph = new Phrase();
				ph.add(System.lineSeparator());
				if (richtextstring instanceof HSSFRichTextString) { // HSSF does not count first run as formatted run when it does not have formatting
					if (richtextstring.getIndexOfFormattingRun(0) != 0) { // index of first formatted run is not at start of text

						String textpart = textstring.substring(0, richtextstring.getIndexOfFormattingRun(0)); // get first run from index 0 to index of first formatted run

						int fontIndex = cell.getCellStyle().getFontIndexAsInt();

						HSSFFont excelFont = workBook.getFontAt(fontIndex);

						Chunk chunk = getChunk(textpart, excelFont, workBook);
						ph.add(chunk);
					} 
				}

				for (int i = 0; i < richtextstring.numFormattingRuns(); i++) { // loop through all formatted runs
					// get index of frormatting run and index of next frormatting run
					int indexofformattingrun = richtextstring.getIndexOfFormattingRun(i);
					int indexofnextformattingrun = textstring.length();
					if ((i+1) < richtextstring.numFormattingRuns()) indexofnextformattingrun = richtextstring.getIndexOfFormattingRun(i+1);
					// formatted text part is the sub string from index of frormatting run to index of next frormatting run
					String textpart = textstring.substring(indexofformattingrun, indexofnextformattingrun);
					// determine used font

					if (richtextstring instanceof HSSFRichTextString) {
						HSSFFont excelFont = null; 
						short fontIndex = ((HSSFRichTextString)richtextstring).getFontOfFormattingRun(i);
						// font index might be HSSFRichTextString.NO_FONT if no formatting is applied to the specified text run
						// then font of the cell should be used.
						if (fontIndex == HSSFRichTextString.NO_FONT) {
							excelFont  = (HSSFFont) workBook.getFontAt(cell.getCellStyle().getFontIndex());   
						} else {
							excelFont = workBook.getFontAt((int)fontIndex);
						} 

						Chunk chunk = getChunk(textpart, excelFont, workBook);
						ph.add(chunk);
					}
				}
				ph.add(System.lineSeparator());
				ph.add(System.lineSeparator());

				//return ph;

			} else {

				int fontIndex = cell.getCellStyle().getFontIndexAsInt();

				HSSFFont excelFont = workBook.getFontAt(fontIndex);

				Chunk chunk = getChunk(sCellVal, excelFont, workBook);

				//Phrase ph = new Phrase();
				ph.add(System.lineSeparator());
				ph.add(chunk);
				ph.add(System.lineSeparator()+System.lineSeparator());

				//return ph;
			} 
			break;

		default:
			int fontIndex = cell.getCellStyle().getFontIndexAsInt();

			HSSFFont excelFont = workBook.getFontAt(fontIndex);

			Chunk chunk = getChunk(sCellVal, excelFont, workBook);

			//Phrase ph = new Phrase();
			ph.add(System.lineSeparator());
			ph.add(chunk);
			ph.add(System.lineSeparator()+System.lineSeparator());
			break;
		}
		
		return ph;
	}
	
	
	private static Chunk getChunk(String sTextpart, HSSFFont excelFont, HSSFWorkbook workBook) {
		
		short sHeight = excelFont.getFontHeight();
		boolean isBold = excelFont.getBold();
		boolean isItalic = excelFont.getItalic();
		
		Font pdfFont = FontFactory.getFont(excelFont.getFontName(), sHeight/HEIGHT_DIVIDE_BY);
	
		if(isBold && isItalic) {
			pdfFont.setStyle(Font.BOLDITALIC);
		}else if(isBold){
			pdfFont.setStyle(Font.BOLD);
			//font.setStyle(Font.NORMAL);
		}else if(isItalic) {
			pdfFont.setStyle(Font.ITALIC);
		}else {
			pdfFont.setStyle(Font.NORMAL);
		}
					
		if(excelFont.getHSSFColor(workBook) != null) {
			short [] rgb = ((HSSFColor)excelFont.getHSSFColor(workBook)).getTriplet();
		
			if(rgb[0] == 255 && rgb[1] == 255 && rgb[2] == 255) {
				rgb[0] = rgb[1] = rgb[2] = 0;
			}
			
			pdfFont.setColor(rgb[0], rgb[1], rgb[2]);			
		} else {
			pdfFont.setColor(0, 0, 0);
		}
		
		return new Chunk(sTextpart, pdfFont);
		
	}
	
	
	private static void setAlignment(Cell cell, PdfPCell pdfCell) {
		CellStyle cellStyle = cell.getCellStyle();
				
		pdfCell.setNoWrap(false);//!cellStyle.getWrapText());
		
		switch (cellStyle.getAlignment()) {
			
			case LEFT:
				pdfCell.setHorizontalAlignment(Paragraph.ALIGN_LEFT);
				break;
			
			case CENTER:
				pdfCell.setHorizontalAlignment(Paragraph.ALIGN_CENTER);
				break;
			
			case RIGHT:
				pdfCell.setHorizontalAlignment(Paragraph.ALIGN_RIGHT);
				break;
				
			case JUSTIFY:
				pdfCell.setHorizontalAlignment(Paragraph.ALIGN_JUSTIFIED);
				
			default:
				pdfCell.setHorizontalAlignment(Paragraph.ALIGN_CENTER);
				break;
		}
	
		
		switch (cellStyle.getVerticalAlignment()) {
			case TOP:
				pdfCell.setVerticalAlignment(Element.ALIGN_TOP);
				break;
			
			case CENTER:
				pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
				break;
			
			case BOTTOM:
				pdfCell.setVerticalAlignment(Element.ALIGN_BOTTOM);
				break;
			
			case JUSTIFY:
				pdfCell.setVerticalAlignment(Element.ALIGN_JUSTIFIED);
				
			default:
				pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
				break;
		}
				
	}
	
	private static void setCellBorder(Cell cell, PdfPCell pdfCell) {
		CellStyle cellStyle = cell.getCellStyle();
				
		pdfCell.setBorderWidthLeft(setBorder(cellStyle.getBorderLeft()));
	    pdfCell.setBorderWidthTop(setBorder(cellStyle.getBorderTop()));
	    pdfCell.setBorderWidthRight(setBorder(cellStyle.getBorderRight()));
	    pdfCell.setBorderWidthBottom(setBorder(cellStyle.getBorderBottom()));
	  
	}
	
	private static float setBorder(BorderStyle borderStyle) {
		switch(borderStyle){
			case NONE:
				return 0f;
				
			case THIN:
				return 0.0001f;
				
			case DOUBLE:
				
				return 0.0001f;
				
			default:
				return 1f;
		}
	}
	
	private static void loadPdfFonts() throws Exception {
		String [] arrXlFonts = {"Cambria", "Calibri", "Helv", "Arial", "Times New Roman", "Bookman Old Style", "Times-Bold", "Tahoma", "Trebuchet MS"};
		
		for(String font : arrXlFonts) {
			mFonts.put(font, Fonts.getPdfFont(font, false, false));
			mFonts.put(font+"b", Fonts.getPdfFont(font, true, false));
			mFonts.put(font+"i", Fonts.getPdfFont(font, false, true));
			mFonts.put(font+"bi", Fonts.getPdfFont(font, true, true));
		}
	}
	
	private static Font getPdfFont(String sFontName, boolean isBold, boolean isItalic) {
		if(isBold && isItalic) {
			return mFonts.get(sFontName+"bi");
		} else if(isBold) {
			return mFonts.get(sFontName+"b");
		} else if(isItalic) {
			return mFonts.get(sFontName+"i");
		} else {
			return mFonts.get(sFontName);
		}
	}
	
	private static Object cloneObject(Object obj){
        try{
            Object clone = obj.getClass().newInstance();
            for (Field field : obj.getClass().getDeclaredFields()) {
                field.setAccessible(true);
                
                field.set(clone, field.get(obj));
            }
            return clone;
        }catch(Exception e){
        	e.printStackTrace();
            return null;
        }
    }
}