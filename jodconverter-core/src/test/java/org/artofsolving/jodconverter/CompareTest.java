package org.artofsolving.jodconverter;

import com.sun.star.document.UpdateDocMode;
import com.sun.star.lang.XComponent;
import java.io.File;
import java.util.HashMap;
import java.util.Map;
import org.apache.commons.io.FilenameUtils;
import org.artofsolving.jodconverter.document.DefaultDocumentFormatRegistry;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.document.DocumentFormatRegistry;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.artofsolving.jodconverter.process.PureJavaProcessManager;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

/**
 *
 * @author lali
 */
public class CompareTest {

    private final File inputFile = new File("/home/lali/Downloads/a.docx");
    private final File expectedFile = new File("/home/lali/Downloads/b.docx");
    private final File outputFile = new File("/home/lali/Downloads/out.docx");
    private final DocumentFormatRegistry formatRegistry = new DefaultDocumentFormatRegistry();
    private OfficeManager officeManager;

    @BeforeTest
    public void startOfficeServer() {
        DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
        configuration.setProcessManager(new PureJavaProcessManager());
        configuration.setOfficeHome("/home/lali/Documents/libreoffice/6.0.4.2/opt/libreoffice6.0");
        configuration.setPortNumber(2004);

        officeManager = configuration.buildOfficeManager();
        officeManager.start();
    }

    @AfterTest
    public void stopOfficeServer() {
        officeManager.stop();
    }

    @Test
    public void testCompare() {
        String inputExtension = FilenameUtils.getExtension(inputFile.getName());
        DocumentFormat inputFormat = formatRegistry.getFormatByExtension(inputExtension);
        
        String outputExtension = FilenameUtils.getExtension(outputFile.getName());
        DocumentFormat outputFormat = formatRegistry.getFormatByExtension(outputExtension);

        CompareTask compareTask = new CompareTask(inputFile, expectedFile, outputFile, inputFormat, outputFormat);
        officeManager.execute(compareTask);
    }

}