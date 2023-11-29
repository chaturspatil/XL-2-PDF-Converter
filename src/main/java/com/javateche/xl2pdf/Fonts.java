package com.javateche.xl2pdf;

import java.io.File;
import java.io.IOException;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.BaseFont;

public class Fonts {
	static String sBasePath = "D:\\My_WORKSSSSSS\\TW-DAE-E2021_Excel-PDF\\fonts";
	
	public static Font getPdfFont(String sXlFontName, boolean isBold, boolean isItalic) throws Exception {
		String sFontFile = "";
	
		switch(sXlFontName) {
		
			case "Cambria":				
				if(isBold && isItalic) {
					sFontFile = "cambriaz.ttf";
				} else if(isBold) {
					sFontFile = "cambriab.ttf";
				} else if(isItalic) {
					sFontFile = "cambriai.ttf";
				} else {
					sFontFile = "Cambria.ttf";
				}
				break;
			
			case "Calibri":
				if(isBold && isItalic) {
					sFontFile = "calibribi.ttf";
				} else if(isBold) {
					sFontFile = "calibrib.ttf";
				} else if(isItalic) {
					sFontFile = "calibrii.ttf";
				} else {
					sFontFile = "calibri.ttf";
				}
				break;
				
			case "Helv":
				if(isBold && isItalic) {
					sFontFile = "Helveticabi.ttf";
				} else if(isBold) {
					sFontFile = "Helveticab.ttf";
				} else if(isItalic) {
					sFontFile = "Helveticai.ttf";
				} else {
					sFontFile = "Helvetica.ttf";
				}
				break;
				
			case "Arial":
				if(isBold && isItalic) {
					sFontFile = "arialbi.ttf";
				} else if(isBold) {
					sFontFile = "arialbd.ttf";
				} else if(isItalic) {
					sFontFile = "ariali.ttf";
				} else {
					sFontFile = "arial.ttf";
				}
				break;
				
			case "Times New Roman":
				if(isBold && isItalic) {
					sFontFile = "timesbi.ttf";
				} else if(isBold) {
					sFontFile = "timesbd.ttf";
				} else if(isItalic) {
					sFontFile = "timesi.ttf";
				} else {
					sFontFile = "times.ttf";
				}
				break;
				
			case "Courier new":
				if(isBold && isItalic) {
					sFontFile = "courbi.ttf";
				} else if(isBold) {
					sFontFile = "courbd.ttf";
				} else if(isItalic) {
					sFontFile = "couri.ttf";
				} else {
					sFontFile = "cour.ttf";
				}
				break;
				
			case "Bookman Old Style":
				if(isBold && isItalic) {
					sFontFile = "BOOKOSBI.TTF";
				} else if(isBold) {
					sFontFile = "BOOKOSB.TTF";
				} else if(isItalic) {
					sFontFile = "BOOKOSI.TTF";
				} else {
					sFontFile = "BOOKOS.TTF";
				}
				break;
				
			case "Trebuchet MS":
				if(isBold && isItalic) {
					sFontFile = "trebucbi.ttf";
				} else if(isBold) {
					sFontFile = "trebucbd.ttf";
				} else if(isItalic) {
					sFontFile = "trebucit.ttf";
				} else {
					sFontFile = "trebuc.ttf";
				}
				break;
				
			case "Tahoma":
				if(isBold && isItalic) {
					sFontFile = "tahomabi.ttf";
				} else if(isBold) {
					sFontFile = "tahomabd.ttf";
				} else if(isItalic) {
					sFontFile = "tahomai.ttf";
				} else {
					sFontFile = "tahoma.ttf";
				}
				break;
				
			case "Times-Bold":
				if(isBold && isItalic) {
					sFontFile = "timesboldbi.ttf";
				} else if(isBold) {
					sFontFile = "timesboldb.ttf";
				} else if(isItalic) {
					sFontFile = "timesboldi.ttf";
				} else {
					sFontFile = "timesbold.ttf";
				}
				break;
				
			default: 
				sXlFontName = "Book Antiqua";
				if(isBold && isItalic) {
					sFontFile = "ANTQUABI.TTF";
				} else if(isBold) {
					sFontFile = "ANTQUAB.TTF";
				} else if(isItalic) {
					sFontFile = "ANTQUAI.TTF";
				} else {
					sFontFile = "BKANT.TTF";
				}
				break;	
		}
		
		BaseFont baseFont = BaseFont.createFont(sBasePath + File.separator + sXlFontName + File.separator + sFontFile, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
		return new Font(baseFont);
	}
	
	
}
