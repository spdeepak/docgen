package com.docgen.features;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.AfterClass;
import org.junit.Assert;
import org.junit.Test;

import com.docgen.converters.ConvertDocxTo;

/**
 * @author deepak.prabhakar
 *
 */
public class DocxXmlMergeImplTest {

    private static final String INPUT_LOCATION = System.getProperty("user.dir") + "/src/test/resources/files/input/";

    private static final String OUTPUT_LOCATION = System.getProperty("user.dir") + "/src/test/resources/files/output/";

    private static final String DOCX_INPUT = INPUT_LOCATION + "docxtemplate.docx";

    private static final String NOTFOUND_INPUT = INPUT_LOCATION + "filnotfound.docx";

    private static final String XML_INPUT = INPUT_LOCATION + "chargesinvoice.xml";

    private static final String DOCX_OUTPUT = OUTPUT_LOCATION + "CustomerOutput.docx";

    private static final String NOTFOUND_DOCX_OUTPUT = OUTPUT_LOCATION + "NotFoundCustomerOutput.docx";

    private static final String PDF_OUTPUT = OUTPUT_LOCATION + "CustomerOutput.pdf";

    private static final String NOTFOUND_PDF_OUTPUT = OUTPUT_LOCATION + "NotFoundCustomerOutput.pdf";

    private static final String HTML_OUTPUT = OUTPUT_LOCATION + "CustomerOutput.html";

    @AfterClass
    public static void deleteFilesAfterCreation() {
        List<String> fileList = new ArrayList<>();
        fileList.add(DOCX_OUTPUT);
        fileList.add(PDF_OUTPUT);
        fileList.add(HTML_OUTPUT);

        for (String files : fileList) {
            if (new File(files).exists()) {
                new File(files).delete();
            }
        }

        if (new File("hf.fo").exists()) {
            Assert.assertTrue(new File("hf.fo").delete());
        }
    }

    private DocxXmlMerge docxXmlMergeImpl = new DocxXmlMerge();

    private DocxUtil docxProperties = new DocxUtil();

    @Test
    public void docxXmlMergeToGetDocxTest() throws IOException {
        docxXmlMergeImpl.toGetDocx(DOCX_INPUT, XML_INPUT, DOCX_OUTPUT);

        Assert.assertTrue(new File(DOCX_OUTPUT).exists());
        Assert.assertEquals(1, docxProperties.getPageCount(DOCX_OUTPUT));

        ConvertDocxTo.html(new FileInputStream(DOCX_OUTPUT), new FileOutputStream(HTML_OUTPUT));
        Assert.assertTrue(new File(HTML_OUTPUT).exists());
    }

    @Test
    public void docxXmlMergeToGetPdfTest() {
        docxXmlMergeImpl.toGetPdf(DOCX_INPUT, XML_INPUT, PDF_OUTPUT);
        Assert.assertTrue(new File(PDF_OUTPUT).exists());
    }

    @Test
    public void exceptionTests() {
        docxXmlMergeImpl.toGetDocx(NOTFOUND_INPUT, XML_INPUT, NOTFOUND_DOCX_OUTPUT);
        Assert.assertTrue(!new File(NOTFOUND_INPUT).exists());
        Assert.assertTrue(!new File(NOTFOUND_DOCX_OUTPUT).exists());

        docxXmlMergeImpl.toGetPdf(NOTFOUND_INPUT, XML_INPUT, NOTFOUND_PDF_OUTPUT);
        Assert.assertTrue(!new File(NOTFOUND_INPUT).exists());
        Assert.assertTrue(!new File(NOTFOUND_PDF_OUTPUT).exists());

    }
}
