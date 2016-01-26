package com.docgen.docx;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.Writer;

public class SubstituteText {

    static class ExtensionFilter implements FilenameFilter {

        String ext;

        public ExtensionFilter(String ext) {
            this.ext = "." + ext;
        }

        public boolean accept(File dir, String name) {
            return name.endsWith(ext);
        }
    }

    public static void main(String[] args) throws Exception {
        final String docxDir = System.getProperty("user.dir") + "/src/test/resources/files/output/"; // directory with
        // the original
        // .docx file
        final String docxName = "CustomerOutput.docx"; // file name of the original .docx
        // file
        final String docxSubName = "Test1Sub.docx"; // file name of the .docx
                                                    // file created with
                                                    // substituted texts

        try {
            // 1. unzip docx file
            ZipUtility.unzip(new File(docxDir, docxName), new File(docxDir, stripFileExt(docxName)));

            // 2a. get list of XML files in /word in unzipped folder
            FilenameFilter ff = new ExtensionFilter("xml");
            File XMLdir = new File(new File(docxDir, stripFileExt(docxName)), "word");
            String[] XMLfiles = XMLdir.list(ff);

            // 2b. read xml files and do text substitution
            if (XMLfiles != null) {
                for (int i = 0; i < XMLfiles.length; i++) {
                    substituteText(new File(XMLdir, XMLfiles[i]), new File(XMLdir, "_" + XMLfiles[i]));
                }
            }

            // 3. zip contents back to docx file
            ZipUtility.zipDirectory(new File(docxDir, stripFileExt(docxName)), new File(docxDir, docxSubName));

            // 4. delete unzipped folder
            cleanDirectory(new File(docxDir, stripFileExt(docxName)));

            // System.out.println("END");
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    /**
     * Function to substitute placeholders with other text IMPORTANT: Placeholders must start and
     * end with %
     */
    private static boolean substituteText(File origFile, File tmpFile) throws Exception {
        StringBuffer sb = new StringBuffer();
        BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(origFile), "UTF-8"));
        Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(tmpFile), "UTF-8"));

        // Names of placeholders, starting and ending with % (to be updated
        // accordingly)
        String placeholder1 = "%Click here to enter text.%";
        String placeholder2 = "%text%";

        // Values to replace placeholders (to be updated accordingly)
        String var1 = escapeHTML("\b"); // %name%
        String var2 = escapeHTML("newText"); // %text%

        String line;
        for (int i = 1; ((line = reader.readLine()) != null); i++) {
            int cursor = 0;
            // Print to file and flush for every 1000 lines
            // To prevent memory error
            if (i % 1000 == 0) {
                writer.write(sb.toString());
                writer.flush();
                sb = new StringBuffer();
            }

            int startIdx = 0;
            int endIdx = 0;
            String result = "";
            while ((startIdx = line.indexOf("%", cursor)) > -1 && (endIdx = line.indexOf("%", startIdx + 1)) > -1) {
                result += line.substring(cursor, startIdx);
                cursor = endIdx + 1;

                String substring = line.substring(startIdx, cursor);
                String stripXML = stripXMLHTMLTags(substring);

                if (stripXML != null && !stripXML.equals("")) {
                    // if is placeholder, replace with text accordingly
                    if (stripXML.equals(placeholder1))
                        result += var1;
                    else if (stripXML.equals(placeholder2))
                        result += var2;
                    else {
                        result += (substring.substring(0, substring.length() - 1));
                        cursor = endIdx;
                    }
                }
            }

            result += line.substring(cursor);
            line = result;

            sb.append(line);
            sb.append("\r\n");
        }
        writer.write(sb.toString());
        writer.flush();
        writer.close();

        reader.close();

        origFile.delete();
        // Rename file (or directory)
        return tmpFile.renameTo(origFile);
    }

    /**
     * Clean directory - delete all files and folders
     */
    private static void cleanDirectory(File d) {
        if (d.exists()) {
            File[] files = d.listFiles();
            for (int i = 0; i < files.length; i++) {
                if (files[i].isDirectory() && files[i].listFiles().length > 0) {
                    cleanDirectory(files[i]);
                }
                files[i].delete();
            }
            d.delete();
        }
    }

    /**
     * Strip string of XML and HTML tags
     */
    private static String stripXMLHTMLTags(String s) {
        return s.replaceAll("\\<.*?>", "");
    }

    /**
     * Strip file name of extension
     */
    private static String stripFileExt(String f) {
        return f.substring(0, f.lastIndexOf("."));
    }

    /**
     * Escape valid html tags
     */
    private static String escapeHTML(String s) {
        if (s == null)
            return s;
        return s.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll("\"", "&quot;")
                .replaceAll("'", "'").replaceAll("/", "/").replaceAll("'", "'");
    }
}
