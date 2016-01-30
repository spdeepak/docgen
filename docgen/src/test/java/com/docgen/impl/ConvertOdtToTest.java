package com.docgen.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.junit.AfterClass;
import org.junit.Test;

import com.docgen.converters.ConvertOdtTo;

public class ConvertOdtToTest {
	
	private static final String INPUT_FOLDER = System.getProperty("user.dir") + "/src/test/resources/files/input/";
	
	private static final String ODT_FILE = INPUT_FOLDER + "document.odt";
	
	private static final String OUTPUT_FOLDER = System.getProperty("user.dir") + "/src/test/resources/files/output/";
	
	private static final String PDF_FILE = OUTPUT_FOLDER + "document.pdf";
	
	@Test
	public void ConverOdtToPdfTest() throws FileNotFoundException{
		ConvertOdtTo.pdf(new FileInputStream(ODT_FILE), new FileOutputStream(PDF_FILE));
	}
	
	@AfterClass
	public void deleteFilesAfterCreation(){
		if(new File(PDF_FILE).exists()){
			new File(PDF_FILE).delete();
		}
	}
}
