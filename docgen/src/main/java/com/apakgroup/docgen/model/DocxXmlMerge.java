package com.apakgroup.docgen.model;

import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * Merges XML data with an appropriate docx template
 * 
 * @author deepak.prabhakar
 *
 */
public interface DocxXmlMerge {

    /**
     * Merges the XML data with the provided docx template to get a docx file with the XML data in
     * it in the format as in the docx template
     * 
     * @param docx_template_location
     *            String location of the template docx file
     * @param xml_data_location
     *            String location of the XML data file
     * @param required_outputfile_name
     *            Required String location for the merged docx file
     */
    public void toGetDocx(String docx_template_location, String xml_data_location, String required_outputfile_name);

    /**
     * @param docx_template_location
     *            Input docx file with extension
     * @param xml_data_location
     *            XML data file with extension
     * @param required_outputfile_name
     *            PDF file name with extension
     */
    public void toGetDocx(FileInputStream docx_template_location, FileInputStream xml_data_location,
            FileOutputStream required_outputfile_name);

    /**
     * Merges the XML data with the provided docx template to get a pdf (only) file with the XML
     * data in it
     * 
     * @param docx_template_location
     *            String location of the template docx file
     * @param xml_data_location
     *            String location of the XML data file
     * @param required_outputfile_name
     *            Required String location for the merged pdf file
     */
    public void toGetPdf(String docx_template_location, String xml_data_location, String required_outputfile_name);

    /**
     * Merges the XML data with the provided docx template to get a pdf file with the XML data in it
     * in the format as in the docx template
     * 
     * @param docx_template
     *            Input docx file
     * @param xml_data
     *            XML data file
     * @param resulting_pdf
     *            PDF file
     */
    public void toGetPdf(FileInputStream docx_template, FileInputStream xml_data, FileOutputStream resulting_pdf);
}
