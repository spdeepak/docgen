package com.docgen.crypt;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.security.GeneralSecurityException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author deepak
 *
 */
public class EncryptDocument {
	
	private static final Logger logger = LoggerFactory.getLogger(EncryptDocument.class);

	public OutputStream docx(InputStream docxFile, String password) {
		try {
			POIFSFileSystem poifsFileSystem = new POIFSFileSystem();
			EncryptionInfo encryptionInfo = new EncryptionInfo(EncryptionMode.agile);
			// EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile,
			// CipherAlgorithm.aes192, HashAlgorithm.sha384, -1, -1, null);

			Encryptor encryptor = encryptionInfo.getEncryptor();
			encryptor.confirmPassword(password);

			OPCPackage opc = OPCPackage.open(docxFile);
			OutputStream os = encryptor.getDataStream(poifsFileSystem);
			opc.save(os);
			opc.close();
			
			FileOutputStream foStream = new FileOutputStream(docxFile.toString());
			poifsFileSystem.writeFilesystem(foStream);
			foStream.close();
			return foStream;
		} catch (FileNotFoundException e) {
			logger.error("Error with the input file: "+e);
		} catch (GeneralSecurityException e) {
			logger.error("Error with getting data Stream for POIFileSystem from Encryptor: "+e);
		} catch (IOException e) {
			logger.error("Error with files: "+e);
		} catch (InvalidFormatException e) {
			logger.error("The file is not of docx format: "+e);
		}
		return null;
	}
}