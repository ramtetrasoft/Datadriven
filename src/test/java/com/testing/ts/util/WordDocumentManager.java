package com.testing.ts.util;

import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordDocumentManager {

String filePath, fileName;
	
	
	/**
	 * Constructor to initialize the Word Document filepath and filename
	 * @param filePath The absolute path where the Word Document is stored
	 * @param fileName The name of the Word Document (without the extension).
	 * Note that .doc files are not supported, only .docx files are supported.
	 */
	public WordDocumentManager(String filePath, String fileName) {
		this.filePath = filePath;
		this.fileName = fileName;
	}
	
	
	/**
	 * Function to create a new Word document
	 */
	public void createDocument() throws Exception{
		
		XWPFDocument document = new XWPFDocument();
		writeIntoFile(document);
	}
		
	private void writeIntoFile(XWPFDocument document) throws Exception{
		String absoluteFilePath = filePath + SuiteUtility.getFileSeparator() +
														fileName + ".docx";
		System.out.println(absoluteFilePath);
		FileOutputStream fileOutputStream;
		try {
			fileOutputStream = new FileOutputStream(absoluteFilePath);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			throw new FrameworkException("The specified file \"" + absoluteFilePath + "\" does not exist!");
		}
		
		try {
			document.write(fileOutputStream);
			fileOutputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
			throw new FrameworkException("Error while writing into the specified Word document \"" + absoluteFilePath + "\"");
		}
	}
	
	/**
	 * Function to add a picture to the Word document
	 * @param pictureFile the picture {@link File} to be inserted
	 */
	/*public void addPicture(File pictureFile) {
		XWPFDocument document = openFileForReading();
		XWPFParagraph paragraph = document.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		
		XWPFRun run = paragraph.createRun();
		run.setText(pictureFile.getName());
		run.addBreak();
		
		try {
			run.addPicture(new FileInputStream(pictureFile),
								XWPFDocument.PICTURE_TYPE_PNG,
								pictureFile.getName(),
								Units.toEMU(200), Units.toEMU(200));
		} catch (InvalidFormatException e) {
			e.printStackTrace();
			throw new FrameworkException("InvalidFormatException thrown while adding a picture file to a Word document");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			throw new FrameworkException("FileNotFoundException thrown while adding a picture file to a Word document");
		} catch (IOException e) {
			e.printStackTrace();
			throw new FrameworkException("IOException thrown while adding a picture file to a Word document");
		}
		
		run.addBreak(BreakType.PAGE);
		
		writeIntoFile(document);
	}*/
	
	/**
	 * Function to add a picture to the Word document
	 * @param pictureFile the picture {@link File} to be inserted
	 * @throws Exception 
	 */
	public void addPicture(File pictureFile) throws Exception {
		CustomXWPFDocument document = openFileForReading();
		XWPFParagraph paragraph = document.createParagraph();
		
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		
		XWPFRun run = paragraph.createRun();
		run.setText(pictureFile.getName());
		//run.addCarriageReturn();
		
		String id;
		try {
			id = document.addPictureData(new FileInputStream(pictureFile), Document.PICTURE_TYPE_PNG);
			
			BufferedImage image = ImageIO.read(pictureFile);
			
			Image tmp = image.getScaledInstance(600, 400, Image.SCALE_SMOOTH);
	        BufferedImage resized = new BufferedImage(600, 400, BufferedImage.TYPE_INT_ARGB);
	        Graphics2D g2d = resized.createGraphics();
	        g2d.drawImage(tmp, 0, 0, null);
	        g2d.dispose();
			document.createPicture(id,
						document.getNextPicNameNumber(Document.PICTURE_TYPE_PNG),
						resized.getWidth(), resized.getHeight());
		} catch (InvalidFormatException | IOException e) {
			e.printStackTrace();
			throw new FrameworkException("Exception thrown while adding a picture file to a Word document");
		}
		
		paragraph = document.createParagraph();
		run = paragraph.createRun();
		run.addBreak(BreakType.PAGE);
		
		writeIntoFile(document);
	}
	
	private CustomXWPFDocument openFileForReading() {
		String absoluteFilePath = filePath + SuiteUtility.getFileSeparator() +
														fileName + ".docx";
		
		FileInputStream fileInputStream;
		try	{
			fileInputStream = new FileInputStream(absoluteFilePath);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			throw new FrameworkException("The specified file \"" + absoluteFilePath + "\" does not exist!");
		}
		
		CustomXWPFDocument document;
		try {
			document = new CustomXWPFDocument(fileInputStream);
		} catch (IOException e) {
			e.printStackTrace();
			throw new FrameworkException("Error while opening the specified Word document \"" + absoluteFilePath + "\"");
		}
		
		return document;
	}
	
	
	
	
}
