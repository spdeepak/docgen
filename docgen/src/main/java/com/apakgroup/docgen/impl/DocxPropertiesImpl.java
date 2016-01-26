package com.apakgroup.docgen.impl;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.apakgroup.docgen.model.DocxProperties;

/**
 * @author deepak.prabhakar
 *
 */
public class DocxPropertiesImpl implements DocxProperties {

    private static final Logger LOGGER = LoggerFactory.getLogger(DocxPropertiesImpl.class);

    @SuppressWarnings("resource")
    @Override
    public int getPageCount(InputStream docx_filepath) throws IOException {
        XWPFDocument document = new XWPFDocument(docx_filepath);
        return document.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();
    }

    @Override
    public int getPageCount(String docx_filepath) throws IOException {
        return getPageCount(new FileInputStream(docx_filepath));
    }

    @Override
    public void removeParagraphOrText(FileInputStream inpudocxfile, String stringToBeReplaced,
            String stringToBeReplacedWith, FileOutputStream outputdocxfile) {
        XWPFDocument document = null;
        try {
            //loading docx file
            document = new XWPFDocument(inpudocxfile);
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    //reading an entire paragraph. So size of list is 1 and index of first element is 0
                    String text = run.getText(0);
                    if (text.contains(stringToBeReplaced)) {
                        text = text.replace(stringToBeReplaced, stringToBeReplacedWith);
                        run.setText(text, 1);
                    }
                }
            }
            document.write(outputdocxfile);
        } catch (Exception e) {
            LOGGER.error("Could not create outputdocxFile --> IOEXception" + e);
        } finally {
            try {
                if (document != null) {
                    document.close();
                }
            } catch (IOException e) {
                LOGGER.error("IOException while closing XWPFDocument in DocxPropertiesImpl.removeParagraph()");
            }
        }
    }

    @Override
    public void removeTextFromtables(FileInputStream inpudocxfile, String stringToBeReplaced,
            String stringToBeReplacedWith, FileOutputStream outputdocxfile) {
        XWPFDocument document = null;
        try {
            //loading docx file
            document = new XWPFDocument(inpudocxfile);
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            for (XWPFRun run : paragraph.getRuns()) {
                                String text = run.getText(0);
                                if (text.contains(stringToBeReplaced)) {
                                    text = text.replace(stringToBeReplaced, stringToBeReplacedWith);
                                    run.setText(text);
                                }
                            }
                        }
                    }
                }
            }
            document.write(new FileOutputStream("output.docx"));
        } catch (FileNotFoundException e) {
            LOGGER.error("FileNotFoundException in DocxPropertiesImpl.removeTextFromtables():" + e);
        } catch (IOException e) {
            LOGGER.error("IOException in DocxPropertiesImpl.removeTextFromtables():" + e);
        } finally {
            try {
                if (document != null) {
                    document.close();
                }
            } catch (IOException e) {
                LOGGER.error("IOException while closing XWPFDocument in DocxPropertiesImpl.removeTextFromtables()");
            }

        }
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        String filename = "CustomerOutput - Copy.docx";
        String folder = System.getProperty("user.dir") + "/src/test/resources/files/output/" + filename;

        XWPFDocument document = new XWPFDocument(new FileInputStream(folder));

        for (XWPFParagraph paragraph : document.getParagraphs()) {
            System.out.println(paragraph.getParagraphText());
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                if (text.contains("Click here to enter text.")) {
                    text = text.replace("Click here to enter text.", "");
                    run.setText(text, 0);
                }
            }
        }

        document.write(new FileOutputStream(folder + "asd.docx"));
    }
}
