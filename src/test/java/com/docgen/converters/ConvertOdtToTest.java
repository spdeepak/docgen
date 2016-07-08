package com.docgen.converters;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.junit.AfterClass;
import org.junit.Test;

public class ConvertOdtToTest {

    private static final String OUTPUT_FOLDER = System.getProperty("user.dir") + "/src/test/resources/files/output/";

    private static final String HTML_FILE = OUTPUT_FOLDER + "document.xhtml";

    private static final String INPUT_FOLDER = System.getProperty("user.dir") + "/src/test/resources/files/input/";

    private static final String ODT_FILE = INPUT_FOLDER + "document.odt";

    private static final String PDF_FILE = OUTPUT_FOLDER + "document.pdf";

    @AfterClass
    public static void deleteFilesAfterCreation() {
        List<String> fileList = new ArrayList<>();
        fileList.add(PDF_FILE);
        fileList.add(HTML_FILE);

        for (String files : fileList) {
            if (new File(files).exists()) {
                new File(files).delete();
            }
        }
        fileList.clear();
    }

    @Test
    public void ConverOdtToHtmlTest() throws FileNotFoundException {
        ConvertOdtTo.html(new FileInputStream(ODT_FILE), new FileOutputStream(PDF_FILE));
    }

    @Test
    public void ConverOdtToPdfTest() throws FileNotFoundException {
        ConvertOdtTo.pdf(new FileInputStream(ODT_FILE), new FileOutputStream(PDF_FILE));
    }
}
