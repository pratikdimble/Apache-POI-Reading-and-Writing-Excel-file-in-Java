package com.eracal.api.employeemanagement.reports;

import java.io.FileOutputStream;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;


public class GenerateOfferLetter {
	public static String logo = "logo-leaf.png";
	public static String output = "D:\\HelloWorld.docx";
	public static void main(String args) {
	try
	{
		
		XWPFDocument document = new XWPFDocument();
		
		// CREATE HEADER AND FOOTER START
		String headerTxt = "Reference: ERACAL/HR/2018";
		String headerText = "PRIVATE & CONFIDENTIAL";
		
		String footerTitle = "ACCEPTANCE";
		String footerNote = "I accept the above mentioned terms and conditions \n";
		String footerData = "Name: ___________________ Date: _______________ Signature: _________________";
		
		XWPFHeaderFooterPolicy headerFooterPolicy = document.getHeaderFooterPolicy();
		  if (headerFooterPolicy == null) headerFooterPolicy = document.createHeaderFooterPolicy();
		  XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
		  XWPFParagraph headerParagraph = header.createParagraph();
		  headerParagraph.setAlignment(ParagraphAlignment.LEFT);
		  XWPFRun headerRun=headerParagraph.createRun();  
		  headerRun.setText(headerTxt);
		  headerRun.setFontFamily("Arial");
		  headerRun.setFontSize(12);
		  
		  XWPFParagraph paragraph = header.createParagraph();
		  paragraph.setAlignment(ParagraphAlignment.CENTER);
		  XWPFRun run=paragraph.createRun(); 
		  run.setText("\n");
		  run.setText(headerText);
		  run.setBold(true);
		  run.setUnderline(UnderlinePatterns.SINGLE);
		  run.setFontFamily("Arial");
		  run.setFontSize(14);
		  
		  XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
		  XWPFParagraph footerParagraph = footer.createParagraph();
		  footerParagraph.setAlignment(ParagraphAlignment.CENTER);
		  XWPFRun footerRun=footerParagraph.createRun();  
		  footerRun.setText(footerTitle);
		  footerRun.setBold(true);
		  footerRun.setFontFamily("Arial");
		  footerRun.setFontSize(9);
		  
		  XWPFParagraph footerParagraph1 = footer.createParagraph();
		  footerParagraph1.setAlignment(ParagraphAlignment.CENTER);
		  XWPFRun footerRun1=footerParagraph1.createRun();  
		  footerRun1.setText(footerNote);
		  footerRun1.setBold(false);
		  footerRun1.setFontFamily("Arial");
		  footerRun1.setFontSize(9);
		  
		  XWPFParagraph footerParagraph2 = footer.createParagraph();
		  footerParagraph2.setAlignment(ParagraphAlignment.CENTER);
		  XWPFRun footerRun2=footerParagraph2.createRun();  
		  footerRun2.setText(footerData);
		  footerRun2.setText("\n");
		  footerRun2.setBold(false);
		  footerRun2.setFontFamily("Arial");
		  footerRun2.setFontSize(9);
		  
		// CREATE HEADER AND FOOTER ENDS
		XWPFParagraph para = document.createParagraph();
		para.setAlignment(ParagraphAlignment.LEFT);
		String string ="Date – 	18th Sep 2018";  
		XWPFRun paraRun = para.createRun();
		paraRun.setText(string);
		paraRun.setFontFamily("Arial");
		paraRun.setFontSize(11);
		paraRun.setBold(false);
		paraRun.setItalic(false);
		
		XWPFParagraph para1 = document.createParagraph();
		para1.setAlignment(ParagraphAlignment.LEFT);
		String string0 ="Mr. Pratik Dimble,";  
		XWPFRun para1Run = para1.createRun();
		para1Run.setText("\n");
		para1Run.setText(string0);
		para1Run.setFontFamily("Arial");
		para1Run.setFontSize(11);
		para1Run.setBold(true);
		para1Run.setItalic(false);
		
		XWPFParagraph para2 = document.createParagraph();
		para2.setAlignment(ParagraphAlignment.LEFT);
		String stringName ="Pune";  
		XWPFRun para2Run = para2.createRun();
		para2Run.setText(stringName);
		para2Run.setFontFamily("Arial");
		para2Run.setFontSize(11);
		para2Run.setBold(true);
		para2Run.setItalic(false);
		
		XWPFParagraph para3 = document.createParagraph();
		para3.setAlignment(ParagraphAlignment.CENTER);
		String stringSub ="Sub: Offer Letter";  
		XWPFRun para3Run = para3.createRun();
		para3Run.setText("\n");
		para3Run.setText(stringSub);
		para3Run.setFontFamily("Arial");
		para3Run.setFontSize(11);
		para3Run.setBold(true );
		para3Run.setItalic(false);
		
		XWPFParagraph para4 = document.createParagraph();
		para4.setAlignment(ParagraphAlignment.LEFT);
		String string4 ="Dear Mr. Pratik Dimble,\n";  
		XWPFRun para4Run = para4.createRun();
		para4Run.setText(string4);
		para4Run.setFontFamily("Arial");
		para4Run.setFontSize(11);
		para4Run.setBold(true);
		para4Run.setItalic(false);
		
		XWPFParagraph para5 = document.createParagraph();
		para5.setAlignment(ParagraphAlignment.BOTH);
		String string5 ="This is with reference to your application and subsequent interview we had, we are pleased to offer you a job at our organization on the following terms and conditions\n";  
		XWPFRun para5Run = para5.createRun();
//		para5Run.setText("\n");
		para5Run.setText(string5);
		para5Run.setFontFamily("Arial");
		para5Run.setFontSize(11);
		para5Run.setBold(false);
		para5Run.setItalic(false);
		
		XWPFParagraph para6 = document.createParagraph();
		para6.setAlignment(ParagraphAlignment.BOTH);
		para6.setStyle("Numbering 1");
		String string6_1="1.";
		String string6_2=" Designation & Place of work: ";
		String string6_3="You will be designated as ";
		String string6_4="“Java Developer” ";
		String string6_5="positioned at ";
		String string6_6="Pune. ";
		String string6_7="The hours will be 48 per week. This position is offered subject to satisfactory reference and pre-employment checks. During your period of job with the company, your services may be posted or transferred to any of the office or division or depts. or units of the company and its group of companies or to any other town or city in India or outside India, without any change in terms and conditions of the offer.";
		para6.setIndentationLeft(720);
		XWPFRun para6Run = para6.createRun();
//		para6Run.setSubscript(VerticalAlign.SUBSCRIPT);
		para6Run.setFontFamily("Arial");
		para6Run.setFontSize(11);
//		para6Run.setText("\n");
		
		para6Run.setText(string6_1);
		para6Run.setBold(false);
		
		para6Run = para6.createRun();
		para6Run.setText(string6_2);
		para6Run.setBold(true);
		
		para6Run = para6.createRun();
		para6Run.setText(string6_3);
		para6Run.setBold(false);
		
		para6Run = para6.createRun();
		para6Run.setText(string6_4);
		para6Run.setBold(true);
		
		para6Run = para6.createRun();
		para6Run.setText(string6_5);
		para6Run.setBold(false);
		
		para6Run = para6.createRun();
		para6Run.setText(string6_6);
		para6Run.setBold(true);
		
		para6Run = para6.createRun();
		para6Run.setText(string6_7);
		para6Run.setBold(false);
//		Numbering 1 Cont.
		
		XWPFParagraph para7 = document.createParagraph();
		para7.setAlignment(ParagraphAlignment.BOTH);
		para7.setStyle("Numbering 1");
//		para7.setNumID(BigInteger.valueOf(2));
//		System.out.println(para7.getNumID());
		String string7_1 = "2. ";
		String string7_2 = "Remuneration: Your remuneration has been finalized with you during the time of final interview. As per our POLICY, You will be getting CTC LPA and you will be kept on observatory period for complete Three Month. After that based on your performance we will offer you an employment. The detailed break up will be provided to you at the time of joining. ";
		String string7_3 = "You will be entitled to 12 days holiday per year pro-rata, plus Company Holidays. The Holiday year runs from Jan 1st - Dec 31st.";
		para7.setIndentationLeft(720);
		XWPFRun para7Run = para7.createRun();
		para7Run.setFontFamily("Arial");
		para7Run.setFontSize(11);
		para7Run.setText(string7_1);
		para7Run.setBold(false);
		
		para7Run = para7.createRun();
		para7Run.setText(string7_2);
		para7Run.setBold(true);
		
		para7Run = para7.createRun();
		para7Run.setText(string7_3);
		para7Run.setBold(false);
		
		XWPFParagraph para8 = document.createParagraph();
		para8.setAlignment(ParagraphAlignment.BOTH);
		para8.setStyle("Numbering 1");
		String string8_1 = "3. ";
		String string8_2 = "Office Timing: ";
		String string8_3 = "Office timing will be 10.00 AM - 8.30 PM. You have complete 48 hours of working in a week. Sat-Sun will be holiday. (Depending upon work load)";
		para8.setIndentationLeft(720);
		XWPFRun para8Run = para8.createRun();
		para8Run.setFontFamily("Arial");
		para8Run.setFontSize(11);
		para8Run.setText(string8_1);
		para8Run.setBold(false);
		
		para8Run = para8.createRun();
		para8Run.setText(string8_2);
		para8Run.setBold(true);
		
		para8Run = para8.createRun();
		para8Run.setText(string8_3);
		para8Run.setBold(false);
		
		XWPFParagraph para9 = document.createParagraph();
		para9.setAlignment(ParagraphAlignment.BOTH);
		String string9_1 ="Reporting: ";
		String string9_2 ="You will report to ";  
		String string9_3 ="Sachin Luniya, BDM : Eracal Software PVT LTD.\n";  
		XWPFRun para9Run = para9.createRun();
		para9Run.setFontFamily("Arial");
		para9Run.setFontSize(11);
//		para9Run.setText("\n");
		para9Run.setText(string9_1);
		para9Run.setBold(true);
		
		para9Run = para9.createRun();
		para9Run.setText(string9_2);
		para9Run.setBold(false);
		
		para9Run = para9.createRun();
		para9Run.setText(string9_3);
		para9Run.setBold(true);
		
		XWPFParagraph para10 = document.createParagraph();
		para10.setAlignment(ParagraphAlignment.BOTH);
		para10.setStyle("Numbering 1");
		String string10_1 = "4. ";
		String string10_2 = "Date of Joining: ";
		String string10_3 = "The above offer stands valid on you joining us immediately, ";
		String string10_4 = "Sep 24th, 2018. ";
		String string10_5 = "You are required to get along with you the following documents on your date of joining for us to enable to issue the appointment letter on your date of joining itself. ";
		para10.setIndentationLeft(720);
		XWPFRun para10Run = para10.createRun();
		para10Run.setFontFamily("Arial");
		para10Run.setFontSize(11);
//		para10Run.setText("\n");
		para10Run.setText(string10_1);
		para10Run.setBold(false);
		
		para10Run = para10.createRun();
		para10Run.setText(string10_2);
		para10Run.setBold(true);
		
		para10Run = para10.createRun();
		para10Run.setText(string10_3);
		para10Run.setBold(false);
		
		para10Run = para10.createRun();
		para10Run.setText(string10_4);
		para10Run.setBold(true);
		
		para10Run = para10.createRun();
		para10Run.setText(string10_5);
		para10Run.setBold(false);
		
		XWPFParagraph para11_1 = document.createParagraph();
		para11_1.setAlignment(ParagraphAlignment.BOTH);
		para11_1.setIndentationLeft(1440);
		String string11_1 = "1. Xerox copies of your educational certificates";
		String string11_2 = "2. 4 I card size photographs";
		String string11_3 = "3. Copy of your ID & address proof";
		String string11_4 = "4. Experience letter, Reliving Letter, Last 3months salary slips.";
		XWPFRun para11_1Run = para11_1.createRun();
		para11_1Run.setFontFamily("Arial");
		para11_1Run.setFontSize(11);
		para11_1Run.setText(string11_1);
		
		XWPFParagraph para11_2 = document.createParagraph();
		para11_2.setAlignment(ParagraphAlignment.BOTH);
		para11_2.setIndentationLeft(1440);
		XWPFRun para11_2Run = para11_2.createRun();
		para11_2Run.setFontFamily("Arial");
		para11_2Run.setFontSize(11);
		para11_2Run.setText(string11_2);
		
		XWPFParagraph para11_3 = document.createParagraph();
		para11_3.setAlignment(ParagraphAlignment.BOTH);
		para11_3.setIndentationLeft(1440);
		XWPFRun para11_3Run = para11_3.createRun();
		para11_3Run.setFontFamily("Arial");
		para11_3Run.setFontSize(11);
		para11_3Run.setText(string11_3);
		
		XWPFParagraph para11_4 = document.createParagraph();
		para11_4.setAlignment(ParagraphAlignment.BOTH);
		para11_4.setIndentationLeft(1440);
		XWPFRun para11_4Run = para11_4.createRun();
		para11_4Run.setFontFamily("Arial");
		para11_4Run.setFontSize(11);
		para11_4Run.setText(string11_4);
		
		
		XWPFParagraph para12 = document.createParagraph();
		para12.setAlignment(ParagraphAlignment.LEFT);
		String string12 ="Thanking you, ";  
		XWPFRun para12Run = para12.createRun();
		para12Run.setFontFamily("Arial");
		para12Run.setFontSize(11);
		para12Run.setText("\n");
		para12Run.setText(string12);
		
		XWPFParagraph para13_1 = document.createParagraph();
		para13_1.setAlignment(ParagraphAlignment.BOTH);
		String string13_1 = "For  Eracal Software PVT LTD ";
		XWPFRun para13_1Run = para13_1.createRun();
		para13_1Run.setFontFamily("Arial");
		para13_1Run.setFontSize(11);
		para13_1Run.setText("\n");
		para13_1Run.setText(string13_1);
		para13_1Run.setBold(true);
		
		XWPFParagraph para13_2 = document.createParagraph();
		para13_2.setAlignment(ParagraphAlignment.BOTH);
		String string13_2 = "Operations – Head ";
		XWPFRun para13_2Run = para13_2.createRun();
		para13_2Run.setFontFamily("Arial");
		para13_2Run.setFontSize(11);
		para13_2Run.setText(string13_2);
		para13_2Run.setBold(true);
		
		FileOutputStream out = new FileOutputStream(output);
		document.write(out);
		out.close();
		document.close();
	}
	catch (Exception e)
		{
			e.printStackTrace();
		}
	}
}
