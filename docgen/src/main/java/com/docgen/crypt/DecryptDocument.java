package com.docgen.crypt;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLConnection;
import java.security.GeneralSecurityException;

import org.apache.commons.io.IOUtils;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author deepak
 *
 */
public class DecryptDocument {

	private static final Logger logger = LoggerFactory.getLogger(DecryptDocument.class);

	/**
	 * @param docxFileStream
	 *            docx File or any XML Based file
	 * @param password
	 *            Password
	 * @return {@link InputStream} of the decrypted file
	 */
	public InputStream docxIn(InputStream docxFileStream, String password) {
		try {
			POIFSFileSystem poifsFileSystem = new POIFSFileSystem(docxFileStream);
			EncryptionInfo encryptionInfo = new EncryptionInfo(poifsFileSystem);
			Decryptor decryptor = Decryptor.getInstance(encryptionInfo);
			if (!decryptor.verifyPassword(password)) {
				logger.error("Wrong Password");
				return null;
			}
			InputStream inputStream = decryptor.getDataStream(poifsFileSystem);
			return inputStream;
		} catch (IOException e) {
			logger.error("Error with the input file" + e);
		} catch (GeneralSecurityException e) {
			logger.error("Wrong Password: " + e);
		}
		return null;
	}

	/**
	 * @param docxFileStream docx File or any XML Based file
	 * @param password Password
	 * @return {@link OutputStream} of the decrypted file
	 */
	public OutputStream docxOut(InputStream docxFileStream, String password) {
		try {
			POIFSFileSystem poifsFileSystem = new POIFSFileSystem(docxFileStream);
			EncryptionInfo encryptionInfo = new EncryptionInfo(poifsFileSystem);
			Decryptor decryptor = Decryptor.getInstance(encryptionInfo);
			if (!decryptor.verifyPassword(password)) {
				logger.error("Wrong Password");
				return null;
			}
			InputStream inputStream = decryptor.getDataStream(poifsFileSystem);
			String extension = URLConnection.guessContentTypeFromStream(inputStream);
			File outputFile = new File(inputStream.toString() + extension);
			IOUtils.copyLarge(inputStream, new FileOutputStream(outputFile));
			inputStream.close();
			extension = null;
			outputFile.deleteOnExit();
			return new FileOutputStream(outputFile);
		} catch (IOException e) {
			logger.error("Error with the input file" + e);
		} catch (GeneralSecurityException e) {
			logger.error("Wrong Password: " + e);
		}
		return null;
	}
}
