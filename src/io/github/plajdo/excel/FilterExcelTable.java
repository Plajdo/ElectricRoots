package io.github.plajdo.excel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.TreeSet;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.Border;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class FilterExcelTable{
	
	static int counter2 = 9;
	static int counterPorc = 1;
	
	public static void main(String[] args) throws Exception{
		create();
	}
	
	public static void create() throws Exception{
		TreeSet<String> strediskaSet = new TreeSet<String>(); 
		
		Workbook kmen = Workbook.getWorkbook(new File("kmen.xls"));
		
		Sheet tabulka = kmen.getSheet(0);
		
		for(int i = 0; i < tabulka.getRows(); i++){
			Cell[] riadok = tabulka.getRow(i);
			
			if(riadok[20].getContents().equals("Šašala")){
				strediskaSet.add(riadok[7].getContents());
				
			}
			
		}
		
		ArrayList<String> strediskaList = new ArrayList<String>(strediskaSet);
		
		strediskaList.forEach((hs) -> {
			try{
				WritableWorkbook output = Workbook.createWorkbook(new File("outputs" + File.separator + "output_" + hs + ".xls"));
				WritableSheet sheet = output.createSheet("Sheet", 0);
				
				Alignment align_left = Alignment.LEFT;
				VerticalAlignment align_top = VerticalAlignment.TOP;
				Alignment align_centre = Alignment.CENTRE;
				VerticalAlignment align_centre_v = VerticalAlignment.CENTRE;
				
				WritableCellFormat thiccFormat = new WritableCellFormat();
				thiccFormat.setFont(new WritableFont(WritableFont.createFont("Calibri"), 11, WritableFont.BOLD));
				thiccFormat.setWrap(true);
				thiccFormat.setAlignment(align_left);
				thiccFormat.setVerticalAlignment(align_top);
				
				WritableCellFormat thinFormat = new WritableCellFormat();
				thinFormat.setFont(new WritableFont(WritableFont.createFont("Calibri"), 11, WritableFont.NO_BOLD));
				
				WritableCellFormat arrayFormat = new WritableCellFormat();
				arrayFormat.setFont(new WritableFont(WritableFont.createFont("Calibri"), 11, WritableFont.NO_BOLD));
				arrayFormat.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.BLACK);
				arrayFormat.setWrap(true);
				arrayFormat.setAlignment(align_left);
				arrayFormat.setVerticalAlignment(align_top);
				
				WritableCellFormat porcFormat = new WritableCellFormat();
				porcFormat.setFont(new WritableFont(WritableFont.createFont("Calibri"), 11, WritableFont.NO_BOLD));
				porcFormat.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.BLACK);
				porcFormat.setWrap(true);
				porcFormat.setAlignment(align_centre);
				porcFormat.setVerticalAlignment(align_centre_v);
				
				WritableCellFormat otherArrayFormat = new WritableCellFormat();
				otherArrayFormat.setFont(new WritableFont(WritableFont.createFont("Calibri"), 11, WritableFont.NO_BOLD));
				otherArrayFormat.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.BLACK);
				otherArrayFormat.setWrap(true);
				otherArrayFormat.setAlignment(align_left);
				otherArrayFormat.setVerticalAlignment(align_centre_v);
				
				sheet.setColumnView(0, 5);
				sheet.setColumnView(1, 10);
				sheet.setColumnView(2, 50);
				sheet.setColumnView(3, 20);
				sheet.setColumnView(4, 20);
				sheet.setColumnView(5, 10);
				sheet.setColumnView(6, 10);
				sheet.setColumnView(7, 10);
				sheet.setColumnView(8, 10);
				sheet.setColumnView(9, 10);
				sheet.setColumnView(10, 10);
				sheet.setColumnView(11, 10);
				sheet.setColumnView(12, 10);
				sheet.setColumnView(13, 10);
				sheet.setColumnView(14, 10);
				sheet.setColumnView(15, 10);
				
				sheet.mergeCells(0, 2, 10, 3);
				sheet.mergeCells(7, 7, 9, 7);
				sheet.mergeCells(10, 7, 11, 7);
				sheet.mergeCells(12, 7, 14, 7);
				
				Label entry01 = new Label(0, 0, "Prevádzka: " + hs, thinFormat);
				Label entry02 = new Label(0, 2, "Protokol o odbornej prehliadke a skúške el. ručného náradia podľa STN 33 1600 a elektrických spotrebičov podľa STN 33 1610 a v zmysle vyh. MPSVaR č.508/2009 Z.z.", thiccFormat);
				Label entry03 = new Label(0, 5, "Vykonaná dňa:", thinFormat);
				Label entry04 = new Label(5, 5, "Merací prístroj:", thinFormat);
				Label entry05 = new Label(9, 5, "Dátum kalibrácie:", thinFormat);
				Label entry06 = new Label(12, 5, "Kalibračný list č.", thinFormat);
				sheet.addCell(entry01);
				sheet.addCell(entry02);
				sheet.addCell(entry03);
				sheet.addCell(entry04);
				sheet.addCell(entry05);
				sheet.addCell(entry06);
				
				Label entry07 = new Label(0, 6, "Prevádzkovateľ: Všeobecná úverová banka a.s., Bratislava, IČO: 313 20 155", thinFormat);
				Label entry08 = new Label(8, 6, "Správca: BK, a.s. Dopravná 19, Piešťany", thinFormat);
				sheet.addCell(entry07);
				sheet.addCell(entry08);
				
				Label entry09 = new Label(0, 7, "Por. číslo", arrayFormat);
				Label entry10 = new Label(1, 7, "Inv. číslo", arrayFormat);
				Label entry11 = new Label(2, 7, "Špecifikácia - Názov, typ", arrayFormat);
				Label entry12 = new Label(3, 7, "Číslo", arrayFormat);
				Label entry13 = new Label(4, 7, "Umiestnenie", arrayFormat);
				Label entry14 = new Label(5, 7, "Skúška chodu", arrayFormat);
				Label entry15 = new Label(6, 7, "Pn (W)      In (A)       Un (V)", arrayFormat);
				Label entry16 = new Label(7, 7, "Skupina - zatriedenie el. spotrebiča, el. mech. náradia", arrayFormat);
				Label entry17 = new Label(10, 7, "Meranie:                           1. Riz - izolačný odpor, 2. Rp - ochran. vodiča", arrayFormat);
				Label entry18 = new Label(12, 7, "Meranie", arrayFormat);
				Label entry19 = new Label(15, 7, "Celkový stav", arrayFormat);
				sheet.addCell(entry09);
				sheet.addCell(entry10);
				sheet.addCell(entry11);
				sheet.addCell(entry12);
				sheet.addCell(entry13);
				sheet.addCell(entry14);
				sheet.addCell(entry15);
				sheet.addCell(entry16);
				sheet.addCell(entry17);
				sheet.addCell(entry18);
				sheet.addCell(entry19);
				
				Label entry20 = new Label(0, 8, "", arrayFormat);
				Label entry21 = new Label(1, 8, "", arrayFormat);
				Label entry22 = new Label(2, 8, "", arrayFormat);
				Label entry23 = new Label(3, 8, "Výr. č., HIM, DHIM, číslo IM", arrayFormat);
				Label entry24 = new Label(4, 8, "", arrayFormat);
				Label entry25 = new Label(5, 8, "", arrayFormat);
				Label entry26 = new Label(6, 8, "", arrayFormat);
				Label entry27 = new Label(7, 8, "Spotrebič", arrayFormat);
				Label entry28 = new Label(8, 8, "Skupina náradia", arrayFormat);
				Label entry29 = new Label(9, 8, "Trieda ochrany", arrayFormat);
				Label entry30 = new Label(10, 8, "Riz (MΩ)", arrayFormat);
				Label entry31 = new Label(11, 8, "Rp (Ω)", arrayFormat);
				Label entry32 = new Label(12, 8, "I - dotykový (mA)", arrayFormat);
				Label entry33 = new Label(13, 8, "I - ochranného vodiča (mA)", arrayFormat);
				Label entry34 = new Label(14, 8, "I - náhrad. unikajúci (mA)", arrayFormat);
				Label entry35 = new Label(15, 8, "", arrayFormat);
				sheet.addCell(entry20);
				sheet.addCell(entry21);
				sheet.addCell(entry22);
				sheet.addCell(entry23);
				sheet.addCell(entry24);
				sheet.addCell(entry25);
				sheet.addCell(entry26);
				sheet.addCell(entry27);
				sheet.addCell(entry28);
				sheet.addCell(entry29);
				sheet.addCell(entry30);
				sheet.addCell(entry31);
				sheet.addCell(entry32);
				sheet.addCell(entry33);
				sheet.addCell(entry34);
				sheet.addCell(entry35);
				
				ArrayList<Polozka> polozkaList = new ArrayList<Polozka>();
				
				for(int j = 0; j < tabulka.getRows(); j++){
					Cell[] riadok = tabulka.getRow(j);
					
					if(riadok[7].getContents().equals(hs)){
						if(riadok[25].getContents().equals("E")){
							polozkaList.add(new Polozka(riadok[1].getContents(), riadok[2].getContents(), riadok[3].getContents(), riadok[10].getContents()));
							
						}
						
					}
					
				}
				
				/*
				 * 10 empty lines
				 */
				for(int k = 0; k < 10; k++){
					polozkaList.add(new Polozka("", "", "", ""));
				}
				
				polozkaList.forEach((polozka) -> {
					Label porc = new Label(0, counter2, String.valueOf(counterPorc), porcFormat);
					Label invc = new Label(1, counter2, polozka.invc, otherArrayFormat);
					Label meno = new Label(2, counter2, polozka.nazov, otherArrayFormat);
					Label vyrc = new Label(3, counter2, polozka.vyrc, otherArrayFormat);
					Label emp1 = new Label(4, counter2, polozka.miesto, otherArrayFormat);
					Label emp2 = new Label(5, counter2, "", otherArrayFormat);
					Label emp3 = new Label(6, counter2, "", otherArrayFormat);
					Label emp4 = new Label(7, counter2, "", otherArrayFormat);
					Label emp5 = new Label(8, counter2, "", otherArrayFormat);
					Label emp6 = new Label(9, counter2, "", otherArrayFormat);
					Label emp7 = new Label(10, counter2, "", otherArrayFormat);
					Label emp8 = new Label(11, counter2, "", otherArrayFormat);
					Label emp9 = new Label(12, counter2, "", otherArrayFormat);
					Label emp10 = new Label(13, counter2, "", otherArrayFormat);
					Label emp11 = new Label(14, counter2, "", otherArrayFormat);
					Label emp12 = new Label(15, counter2, "", otherArrayFormat);
					
					try{
						sheet.addCell(porc);
						sheet.addCell(invc);
						sheet.addCell(meno);
						sheet.addCell(vyrc);
						sheet.addCell(emp1);
						sheet.addCell(emp2);
						sheet.addCell(emp3);
						sheet.addCell(emp4);
						sheet.addCell(emp5);
						sheet.addCell(emp6);
						sheet.addCell(emp7);
						sheet.addCell(emp8);
						sheet.addCell(emp9);
						sheet.addCell(emp10);
						sheet.addCell(emp11);
						sheet.addCell(emp12);
					}catch(WriteException e){
						e.printStackTrace();
					}
					
					counter2++;
					counterPorc++;
					
				});
				
				for(int l = 0; l < counter2; l++){
					if(l == 7 || l == 8){
						continue;
					}
					sheet.setRowView(l, 15 * 20);
				}
				
				counter2 = 9;
				counterPorc = 1;
				
				output.write();
				output.close();
				
			}catch(IOException | WriteException e){
				e.printStackTrace();
			}
		});
		
		kmen.close();
		
	}
	
}

class Polozka{
	
	public String invc, nazov, vyrc, miesto;
	public int id;
	
	public Polozka(String invc, String nazov, String vyrc, String miesto){
		this.invc = invc;
		this.nazov = nazov;
		this.vyrc = vyrc;
		this.miesto = miesto;
	}
	
}
