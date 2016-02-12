
package com.docgen.converters;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Used to convert docx files to other formats
 * 
 * @author deepak
 */
public class ConvertDocxTo {

	private static final Logger logger = LoggerFactory.getLogger(ConvertDocxTo.class);

	/**
	 * Converts docx to pdf<br/>
	 * <i><u>File path should include the extensions of the file as well</u></i>
	 * 
	 * @param inputdocxfilepath
	 *            File location fo the docx file
	 * @param outputpdffilepath
	 *            File location of the pdf file
	 */
	public static OutputStream pdf(FileInputStream inputdocxfilepath, FileOutputStream outputpdffilepath) {
		XWPFDocument docxFile = null;
		PdfOptions pdfOptions = null;
		try {
			docxFile = new XWPFDocument(inputdocxfilepath);
			pdfOptions = PdfOptions.create();
			PdfConverter.getInstance().convert(docxFile, outputpdffilepath, pdfOptions);
			return outputpdffilepath;
		} catch (IOException e) {
			logger.error("IOException in ConvertDocxTo.pdf(): " + e);
		} finally {
			try {
				if (docxFile != null) {
					docxFile.close();
				}
				if (pdfOptions != null) {
					pdfOptions = null;
				}
			} catch (IOException e) {
				logger.error("IOException while closing docx file in ConvertDocxTo.pdf(): " + e);
			}
		}
		return null;
	}

	/**
	 * Converts docx to html <br/>
	 * <i><u>File path should include the extensions of the file as well</u></i>
	 * 
	 * @param inputDocxFilePath
	 *            locaiton of the docx file which is to be converted to html
	 * @param outputHtmlFilePath
	 *            location of the html file
	 */
	public static OutputStream html(FileInputStream inputDocxFilePath, FileOutputStream outputHtmlFilePath) {
		XWPFDocument xwpfDocument = null;
		XHTMLOptions xhtmlOptions = null;
		try {
			xwpfDocument = new XWPFDocument(inputDocxFilePath);
			xhtmlOptions = XHTMLOptions.create();// .URIResolver(new
													// FileURIResolver(new
													// File("word/media")))
			XHTMLConverter.getInstance().convert(xwpfDocument, outputHtmlFilePath, xhtmlOptions);
			return outputHtmlFilePath;
		} catch (IOException e) {
			logger.error("IOException in ConvertDocxTo.html(): " + e);
		} finally {
			try {
				if (xwpfDocument != null) {
					xwpfDocument.close();
				}
				if (xhtmlOptions != null) {
					xhtmlOptions = null;
				}
			} catch (IOException e) {
				logger.error("IOException while closing the xwpfDocument in ConvertDocxTo.html(): " + e);
			}

		}
		return null;
	}

}
