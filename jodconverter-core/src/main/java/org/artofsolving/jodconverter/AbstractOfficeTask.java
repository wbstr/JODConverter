//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter;

import static org.artofsolving.jodconverter.office.OfficeUtils.SERVICE_DESKTOP;
import static org.artofsolving.jodconverter.office.OfficeUtils.cast;
import static org.artofsolving.jodconverter.office.OfficeUtils.toUnoProperties;
import static org.artofsolving.jodconverter.office.OfficeUtils.toUrl;

import java.io.File;
import java.util.Map;

import org.artofsolving.jodconverter.office.OfficeContext;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeTask;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.task.ErrorCodeIOException;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;
import java.util.HashMap;
import org.artofsolving.jodconverter.document.DocumentFamily;
import org.artofsolving.jodconverter.document.DocumentFormat;

public abstract class AbstractOfficeTask implements OfficeTask {

    private final File inputFile;
    private final File outputFile;
    private final DocumentFormat inputFormat;
    private final DocumentFormat outputFormat;
    private final Map<String, Object> loadProperties = new HashMap<String, Object>();
    private final Map<String, Object> storeProperties = new HashMap<String, Object>();

    public AbstractOfficeTask(File inputFile, File outputFile, DocumentFormat inputFormat, DocumentFormat outputFormat) {
        this.inputFile = inputFile;
        this.outputFile = outputFile;
        this.inputFormat = inputFormat;
        this.outputFormat = outputFormat;
    }

    @Override
    public void execute(OfficeContext context) throws OfficeException {
        XComponent document = null;
        try {
            document = loadDocument(context, inputFile);
            modifyDocument(document);
            storeDocument(document, outputFile);
            closeDocument(document);
        } catch (OfficeException officeException) {
            throw officeException;
        } catch (Exception exception) {
            throw new OfficeException("conversion failed", exception);
        } finally {
            if (document != null) {
                XCloseable closeable = cast(XCloseable.class, document);
                if (closeable != null) {
                    try {
                        closeable.close(true);
                    } catch (CloseVetoException closeVetoException) {
                        // whoever raised the veto should close the document
                    }
                } else {
                    document.dispose();
                }
            }
        }
    }

    public void addLoadProperty(String key, Object value) {
        loadProperties.put(key, value);
    }

    public void addStoreProperty(String key, Object value) {
        storeProperties.put(key, value);
    }

    private XComponent loadDocument(OfficeContext context, File inputFile) throws OfficeException {
        if (!inputFile.exists()) {
            throw new OfficeException("input document not found");
        }

        Map<String, ?> inputFormatProperties = inputFormat.getLoadProperties();
        if (inputFormatProperties != null) {
            loadProperties.putAll(inputFormatProperties);
        }

        XComponent document = null;
        XComponentLoader loader = cast(XComponentLoader.class, context.getService(SERVICE_DESKTOP));
        try {
            document = loader.loadComponentFromURL(toUrl(inputFile), "_blank", 0, toUnoProperties(loadProperties));
        } catch (IllegalArgumentException illegalArgumentException) {
            throw new OfficeException("could not load document: " + inputFile.getName(), illegalArgumentException);
        } catch (ErrorCodeIOException errorCodeIOException) {
            throw new OfficeException("could not load document: " + inputFile.getName() + "; errorCode: " + errorCodeIOException.ErrCode, errorCodeIOException);
        } catch (IOException ioException) {
            throw new OfficeException("could not load document: " + inputFile.getName(), ioException);
        }
        if (document == null) {
            throw new OfficeException("could not load document: " + inputFile.getName());
        }
        return document;
    }

    /**
     * Override to modify the document after it has been loaded and before it
     * gets saved in the new format.
     * <p>
     * Does nothing by default.
     *
     * @param document
     * @throws OfficeException
     */
    protected abstract void modifyDocument(XComponent document) throws OfficeException;

    private void storeDocument(XComponent document, File outputFile) throws OfficeException {
        DocumentFamily family = OfficeDocumentUtils.getDocumentFamily(document);
        Map<String, ?> outputFormatProperties = outputFormat.getStoreProperties(family);
        if (outputFormatProperties != null) {
            storeProperties.putAll(outputFormatProperties);
        }

        try {
            cast(XStorable.class, document).storeToURL(toUrl(outputFile), toUnoProperties(storeProperties));
        } catch (ErrorCodeIOException errorCodeIOException) {
            throw new OfficeException("could not store document: " + outputFile.getName() + "; errorCode: " + errorCodeIOException.ErrCode, errorCodeIOException);
        } catch (IOException ioException) {
            throw new OfficeException("could not store document: " + outputFile.getName(), ioException);
        }
    }

    private void closeDocument(XComponent document) throws OfficeException {
        XCloseable xCloseable = UnoRuntime.queryInterface(XCloseable.class, document);

        if (xCloseable != null) {
            try {
                xCloseable.close(false);
            } catch (CloseVetoException ex) {
                throw new OfficeException("could not close document: " + inputFile.getName(), ex);
            }
        } else {
            XComponent xComp
                    = UnoRuntime.queryInterface(XComponent.class, document);

            xComp.dispose();
        }
    }

}
