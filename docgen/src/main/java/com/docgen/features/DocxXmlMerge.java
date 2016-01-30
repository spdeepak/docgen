package com.docgen.features;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.docgen.converters.ConvertDocxTo;

/**
 * Merges XML data with the provided docx tmeplate.
 * 
 * @author deepak.prabhakar
 *
 */
public class DocxXmlMerge {

    private static final Logger LOGGER = LoggerFactory.getLogger(DocxXmlMerge.class);

    public void toGetDocx(FileInputStream docx_template_location, FileInputStream xml_data_location,
            FileOutputStream required_outputfile_name) {
        try {
            //Loads docx Template File
            WordprocessingMLPackage wordprocessingMLPackage = Docx4J.load(docx_template_location);
            // Do the binding:
            Docx4J.bind(wordprocessingMLPackage, xml_data_location, Docx4J.FLAG_BIND_INSERT_XML
                    & Docx4J.FLAG_BIND_REMOVE_SDT);
            Docx4J.save(wordprocessingMLPackage, required_outputfile_name, Docx4J.FLAG_NONE);
        } catch (Docx4JException e) {
            LOGGER.error("Docx4j Library Exception: " + e);
        }
    }

    public void toGetDocx(String docx_template_location, String xml_data_location, String required_outputfile_name) {
        try {
            toGetDocx(new FileInputStream(docx_template_location), new FileInputStream(xml_data_location),
                    new FileOutputStream(required_outputfile_name));
        } catch (FileNotFoundException e) {
            LOGGER.error("This file is not present: " + e);
        }

    }

    public void toGetPdf(FileInputStream docx_template, FileInputStream xml_data, FileOutputStream resulting_pdf) {
        try {
            //Load Docx Template
            WordprocessingMLPackage wordprocessingMLPackage = Docx4J.load(docx_template);
            //Load XML Data into the above Docx Template
            Docx4J.bind(wordprocessingMLPackage, xml_data, Docx4J.FLAG_NONE);
            //Temporary docx file
            File temp_doc_file = new File(System.getProperty("java.io.tmpdir") + "tempmerged.docx");
            //save the docx output with all the xml data in it
            Docx4J.save(wordprocessingMLPackage, new FileOutputStream(temp_doc_file), Docx4J.FLAG_NONE);
            //convert the docx file to pdf
            ConvertDocxTo.pdf(new FileInputStream(temp_doc_file), resulting_pdf);

            if (temp_doc_file.exists()) {
                temp_doc_file.delete();
            }

        } catch (Docx4JException e) {
            LOGGER.error("Docx4JException: " + e);
        } catch (FileNotFoundException e) {
            LOGGER.error("This file is not present: " + e);
        }
    }

    public void toGetPdf(String docx_template_location, String xml_data_location, String required_outputfile_name) {
        try {
            toGetPdf(new FileInputStream(docx_template_location), new FileInputStream(xml_data_location),
                    new FileOutputStream(required_outputfile_name));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
}
