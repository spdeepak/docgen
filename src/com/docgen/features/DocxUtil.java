package com.docgen.features;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Deepak
 *
 */
public class DocxUtil {

	private static final Logger logger = LoggerFactory.getLogger(DocxUtil.class);

	@SuppressWarnings("resource")
	public int getPageCount(InputStream docx_filepath) throws IOException {
		return new XWPFDocument(docx_filepath).getProperties().getExtendedProperties().getUnderlyingProperties()
				.getPages();
	}

	public int getPageCount(String docx_filepath) throws IOException {
		return getPageCount(new FileInputStream(docx_filepath));
	}

	public OutputStream removeTextFromDocx(FileInputStream inpudocxfile, String stringToBeReplaced,
			String stringToBeReplacedWith, FileOutputStream outputdocxfile) {
		XWPFDocument document = null;
		try {
			// loading docx file
			document = new XWPFDocument(inpudocxfile);
			for (XWPFParagraph paragraph : document.getParagraphs()) {
				List<XWPFRun> runs = paragraph.getRuns();
				for (XWPFRun run : runs) {
					// reading an entire paragraph. So size of list is 1 and
					// index of first element is 0
					String text = run.getText(0);
					if (text != null) {
						if (text.contains(stringToBeReplaced)) {
							text = text.replace(stringToBeReplaced, stringToBeReplacedWith);
							text = text.trim();
							run.setText(text, 0);
						}
					}
				}
			}
			for (XWPFTable table : document.getTables()) {
				for (XWPFTableRow row : table.getRows()) {
					for (XWPFTableCell cell : row.getTableCells()) {
						for (XWPFParagraph paragraph : cell.getParagraphs()) {
							for (XWPFRun run : paragraph.getRuns()) {
								String text = run.getText(0);
								if (text != null) {
									if (text.contains(stringToBeReplaced)) {
										text = text.replace(stringToBeReplaced, stringToBeReplacedWith);
										text = text.trim();
										run.setText(text, 0);
									}
								}
							}
						}
					}
				}
			}
			document.write(outputdocxfile);
			document.close();
			return outputdocxfile;
		} catch (IOException e) {
			logger.error("IOException in DocxPropertiesImpl.removeTextFromDocx():" + e);
		}
		return null;
	}
}
