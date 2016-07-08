package com.docgen.converters;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.odftoolkit.odfdom.converter.pdf.PdfConverter;
import org.odftoolkit.odfdom.converter.pdf.PdfOptions;
import org.odftoolkit.odfdom.converter.xhtml.XHTMLConverter;
import org.odftoolkit.odfdom.converter.xhtml.XHTMLOptions;
import org.odftoolkit.odfdom.doc.OdfTextDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Deepak
 *
 */
public class ConvertOdtTo {

    private static final Logger logger = LoggerFactory.getLogger(ConvertOdtTo.class);

    public static OutputStream pdf(FileInputStream inputOdtFilePath, FileOutputStream outputPdfFilePath) {
        try {
            OdfTextDocument odtFile = OdfTextDocument.loadDocument(inputOdtFilePath);
            PdfOptions pdfOptions = PdfOptions.create();
            PdfConverter.getInstance().convert(odtFile, outputPdfFilePath, pdfOptions);
            return outputPdfFilePath;
        } catch (Exception e) {
            logger.error("Excpetion while convertinf ODT to PDF: " + e);
        }
        return null;
    }

    public static OutputStream html(FileInputStream inputOdtFilePath, FileOutputStream outputHtmlFilePath) {
        try {
            OdfTextDocument odtFile = OdfTextDocument.loadDocument(inputOdtFilePath);
            XHTMLOptions htmlOptions = XHTMLOptions.create();
            XHTMLConverter.getInstance().convert(odtFile, outputHtmlFilePath, htmlOptions);
            return outputHtmlFilePath;
        } catch (Exception e) {
            logger.error("Excpetion while convertinf ODT to HTML: " + e);
        }
        return null;
    }

}
