
package com.docgen.converters;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.odftoolkit.odfdom.converter.pdf.PdfConverter;
import org.odftoolkit.odfdom.converter.pdf.PdfOptions;
import org.odftoolkit.odfdom.doc.OdfTextDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author deepak
 */
public class ConvertOdtTo {

	private static final Logger LOGGER = LoggerFactory.getLogger(ConvertOdtTo.class);

	public static void pdf(FileInputStream inputodtfilepath, FileOutputStream outputpdffilepath) {
		try {
			OdfTextDocument odtFile = OdfTextDocument.loadDocument(inputodtfilepath);
			PdfOptions pdfOptions = PdfOptions.create();
			PdfConverter.getInstance().convert(odtFile, outputpdffilepath, pdfOptions);
		} catch (Exception e) {
			LOGGER.error("Excpetion while convertinf ODT to PDF: " + e);
		}

	}

}
