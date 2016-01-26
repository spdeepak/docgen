package com.docgen.impl;

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
import com.docgen.model.DocxXmlMerge;

/**
 * Merges XML data with the provided docx tmeplate.
 * 
 * @author deepak.prabhakar
 *
 */
public class DocxXmlMergeImpl implements DocxXmlMerge {

    private static final Logger LOGGER = LoggerFactory.getLogger(DocxXmlMergeImpl.class);

    @Override
    public void toGetDocx(FileInputStream docx_template_location, FileInputStream xml_data_location,
            FileOutputStream required_outputfile_name) {
        try {
            //Loads docx Template File
            WordprocessingMLPackage wordprocessingMLPackage = Docx4J.load(docx_template_location);

            // Do the binding:
            // FLAG_NONE means that all the steps of the binding will be done,
            // otherwise you could pass a combination of the following flags:
            // FLAG_BIND_INSERT_XML: inject the passed XML into the document
            // FLAG_BIND_BIND_XML: bind the document and the xml (including any OpenDope handling)
            // FLAG_BIND_REMOVE_SDT: remove the content controls from the document (only the content remains)
            // FLAG_BIND_REMOVE_XML: remove the custom xml parts from the document 

            //Docx4J.bind(wordMLPackage, xmlStream, Docx4J.FLAG_NONE);
            //If a document doesn't include the Opendope definitions, eg. the XPathPart,
            //then the only thing you can do is insert the xml
            //the example document binding-simple.docx doesn't have an XPathPart....
            Docx4J.bind(wordprocessingMLPackage, xml_data_location, Docx4J.FLAG_BIND_INSERT_XML
                    & Docx4J.FLAG_BIND_REMOVE_SDT);
            Docx4J.save(wordprocessingMLPackage, required_outputfile_name, Docx4J.FLAG_NONE);

        } catch (Docx4JException e) {
            LOGGER.error("Docx4j Library Exception: " + e);
        }
    }

    @Override
    public void toGetDocx(String docx_template_location, String xml_data_location, String required_outputfile_name) {
        try {
            toGetDocx(new FileInputStream(docx_template_location), new FileInputStream(xml_data_location),
                    new FileOutputStream(required_outputfile_name));

        } catch (FileNotFoundException e) {
            LOGGER.error("This file is not present: " + e);
        }

    }

    @Override
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

    @Override
    public void toGetPdf(String docx_template_location, String xml_data_location, String required_outputfile_name) {
        try {
            toGetPdf(new FileInputStream(docx_template_location), new FileInputStream(xml_data_location),
                    new FileOutputStream(required_outputfile_name));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
}
