package com.apakgroup.docgen.model;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Get Properties of a docx file
 * 
 * @author deepak.prabhakar
 *
 */
public interface DocxProperties {

    /**
     * Get the number of pages present in a docx file
     * 
     * @param docx_filepath
     *            {@link InputStream} of the docx file
     * @return number of pages in the docx file
     * @throws IOException
     */
    public int getPageCount(InputStream docx_filepath) throws IOException;

    /**
     * Get the number of pages present in a docx file
     * 
     * @param docx_filepath
     *            String path of the docx file
     * @return number of pages in the docx file
     * @throws IOException
     */
    public int getPageCount(String docx_filepath) throws IOException;

    /**
     * 
     * Replace entire paragraph or text in a docx file
     * 
     * @param inpudocxfile
     *            Input docx file in which a paragrpah or text is supposed to be replaced with
     *            something else
     * @param stringToBeReplaced
     *            String which is to be replaced
     * @param stringToBeReplacedWith
     *            String which will be replaced with
     * @param outputdocxfile
     *            Output docx file with all the replaced paragraphs/text
     */
    public void removeParagraphOrText(FileInputStream inpudocxfile, String stringToBeReplaced,
            String stringToBeReplacedWith, FileOutputStream outputdocxfile);

    /**
     * 
     * Replace text from a table in a docx file
     * 
     * @param inpudocxfile
     *            Input docx file in which a paragrpah or text inside a table is supposed to be
     *            replaced with something else
     * @param stringToBeReplaced
     *            String which is to be replaced
     * @param stringToBeReplacedWith
     *            String which will be replaced with
     * @param outputdocxfile
     *            Output docx file with all the replaced paragraphs/text inside tables
     */
    public void removeTextFromtables(FileInputStream inpudocxfile, String stringToBeReplaced,
            String stringToBeReplacedWith, FileOutputStream outputdocxfile);
}
