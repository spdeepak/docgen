package com.apakgroup.docgen.converters;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

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
 * @author deepak.prabhakar
 *
 */
public class ConvertDocxTo {

    private static final Logger LOGGER = LoggerFactory.getLogger(ConvertDocxTo.class);

    /**
     * Converts docx to pdf<br/>
     * <i><u>File path should include the extensions of the file as well</u></i>
     * 
     * @param inputdocxfilepath
     *            File location fo the docx file
     * 
     * @param outputpdffilepath
     *            File location of the pdf file
     * 
     */
    public static void pdf(FileInputStream inputdocxfilepath, FileOutputStream outputpdffilepath) {
        XWPFDocument docxFile = null;
        PdfOptions pdfOptions = null;
        try {
            docxFile = new XWPFDocument(inputdocxfilepath);
            pdfOptions = PdfOptions.create();
            PdfConverter.getInstance().convert(docxFile, outputpdffilepath, pdfOptions);
        } catch (IOException e) {
            LOGGER.error("IOException in ConvertDocxTo.pdf(): " + e);
        } finally {
            try {
                if (docxFile != null) {
                    docxFile.close();
                }
                if (pdfOptions != null) {
                    pdfOptions = null;
                }
            } catch (IOException e) {
                LOGGER.error("IOException while closing docx file in ConvertDocxTo.pdf(): " + e);
            }
        }
    }

    /**
     * 
     * Converts docx to html <br/>
     * <i><u>File path should include the extensions of the file as well</u></i>
     * 
     * @param inputdocxfilepath
     *            locaiton of the docx file which is to be converted to html
     * @param outputpdffilepath
     *            location of the html file
     */
    public static void html(FileInputStream inputdocxfilepath, FileOutputStream outputpdffilepath) {
        XWPFDocument xwpfDocument = null;
        XHTMLOptions xhtmlOptions = null;
        try {
            xwpfDocument = new XWPFDocument(inputdocxfilepath);
            xhtmlOptions = XHTMLOptions.create();//.URIResolver(new FileURIResolver(new File("word/media")))
            XHTMLConverter.getInstance().convert(xwpfDocument, outputpdffilepath, xhtmlOptions);
        } catch (IOException e) {
            LOGGER.error("IOException in ConvertDocxTo.html(): " + e);
        } finally {
            try {
                if (xwpfDocument != null) {
                    xwpfDocument.close();
                }
                if (xhtmlOptions != null) {
                    xhtmlOptions = null;
                }
            } catch (IOException e) {
                LOGGER.error("IOException while closing the xwpfDocument in ConvertDocxTo.html(): " + e);
            }

        }
    }

}
