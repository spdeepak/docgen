package com.docgen.features;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.docgen.converters.ConvertDocxTo;

/**
 * Merges XML data with the provided docx template.
 * 
 * @author deepak.prabhakar
 *
 */
public class DocxXmlMerge {

    private static final Logger logger = LoggerFactory.getLogger(DocxXmlMerge.class);

    /**
     * @param docx_template_location
     * @param xml_data_location
     * @param required_outputfile_name
     * @return {@link OutputStream} of merged file
     */
    public OutputStream toGetDocx(FileInputStream docx_template_location, FileInputStream xml_data_location,
            FileOutputStream required_outputfile_name) {
        try {
            //Loads docx Template File
            WordprocessingMLPackage wordprocessingMLPackage = Docx4J.load(docx_template_location);
            // Do the binding:
            Docx4J.bind(wordprocessingMLPackage, xml_data_location, Docx4J.FLAG_BIND_INSERT_XML
                    & Docx4J.FLAG_BIND_REMOVE_SDT);
            Docx4J.save(wordprocessingMLPackage, required_outputfile_name, Docx4J.FLAG_NONE);
            return required_outputfile_name;
        } catch (Docx4JException e) {
            logger.error("Docx4j Library Exception: " + e);
        }
        return null;
    }

    /**
     * @param docx_template_location
     * @param xml_data_location
     * @param required_outputfile_name
     * @return {@link OutputStream} of merged file
     */
    public OutputStream toGetDocx(String docx_template_location, String xml_data_location, String required_outputfile_name) {
        return toGetDocx(new File(docx_template_location), new File(xml_data_location),
		        new File(required_outputfile_name));
    }
    
    /**
     * @param docx_template_location
     * @param xml_data_location
     * @param required_outputfile_name
     * @return {@link OutputStream} of merged file
     */
    public OutputStream toGetDocx(File docx_template_location, File xml_data_location, File required_outputfile_name) {
        try {
            return toGetDocx(new FileInputStream(docx_template_location), new FileInputStream(xml_data_location),
                    new FileOutputStream(required_outputfile_name));
        } catch (FileNotFoundException e) {
            logger.error("This file is not present: " + e);
        }
        return null;
    }

    /**
     * @param docx_template
     * @param xml_data
     * @param resulting_pdf
     * @return {@link OutputStream} of merged file
     */
    public OutputStream toGetPdf(FileInputStream docx_template, FileInputStream xml_data, FileOutputStream resulting_pdf) {
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
            return ConvertDocxTo.pdf(new FileInputStream(temp_doc_file), resulting_pdf);
        } catch (Docx4JException e) {
            logger.error("Docx4JException: " + e);
        } catch (FileNotFoundException e) {
            logger.error("This file is not present: " + e);
        }
        return null;
    }

    /**
     * @param docx_template_location
     * @param xml_data_location
     * @param required_outputfile_name
     * @return {@link OutputStream} of merged file
     */
    public OutputStream toGetPdf(String docx_template_location, String xml_data_location, String required_outputfile_name) {
        try {
            return toGetPdf(new FileInputStream(docx_template_location), new FileInputStream(xml_data_location),
                    new FileOutputStream(required_outputfile_name));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return null;
    }
}
