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

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.io.FilenameUtils;
import org.artofsolving.jodconverter.document.DefaultDocumentFormatRegistry;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.document.DocumentFormatRegistry;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeManager;

import com.sun.star.document.UpdateDocMode;

public class OfficeDocumentConverter {

    private final OfficeManager officeManager;
    private final DocumentFormatRegistry formatRegistry;

    private Map<String, ?> defaultLoadProperties = createDefaultLoadProperties();
    private Map<String, ?> defaultStoreProperties = createDefaultStoreProperties();

    public OfficeDocumentConverter(OfficeManager officeManager) {
        this(officeManager, new DefaultDocumentFormatRegistry());
    }

    public OfficeDocumentConverter(OfficeManager officeManager, DocumentFormatRegistry formatRegistry) {
        this.officeManager = officeManager;
        this.formatRegistry = formatRegistry;
    }

    private Map<String, Object> createDefaultLoadProperties() {
        Map<String, Object> loadProperties = new HashMap<String, Object>();
        loadProperties.put("Hidden", true);
        loadProperties.put("ReadOnly", true);
        loadProperties.put("UpdateDocMode", UpdateDocMode.QUIET_UPDATE);
        return loadProperties;
    }

    private Map<String, Object> createDefaultStoreProperties() {
        Map<String, Object> loadProperties = new HashMap<String, Object>();
        loadProperties.put("Overwrite", true);
        return loadProperties;
    }

    public void setDefaultLoadProperties(Map<String, ?> defaultLoadProperties) {
        this.defaultLoadProperties = defaultLoadProperties;
    }

    public void setDefaultStoreProperties(Map<String, ?> defaultStoreProperties) {
        this.defaultStoreProperties = defaultStoreProperties;
    }

    public DocumentFormatRegistry getFormatRegistry() {
        return formatRegistry;
    }

    public void convert(File inputFile, File outputFile) throws OfficeException {
        String inputExtension = FilenameUtils.getExtension(inputFile.getName());
        DocumentFormat inputFormat = formatRegistry.getFormatByExtension(inputExtension);

        String outputExtension = FilenameUtils.getExtension(outputFile.getName());
        DocumentFormat outputFormat = formatRegistry.getFormatByExtension(outputExtension);

        convert(inputFile, outputFile, inputFormat, outputFormat);
    }

    public void convert(File inputFile, File outputFile, DocumentFormat inputFormat, DocumentFormat outputFormat) throws OfficeException {
        StandardConversionTask conversionTask = new StandardConversionTask(inputFile, outputFile, inputFormat, outputFormat);

        for (Map.Entry<String, ? extends Object> entry : defaultLoadProperties.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();

            conversionTask.addLoadProperty(key, value);
        }

        for (Map.Entry<String, ? extends Object> entry : defaultStoreProperties.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();

            conversionTask.addStoreProperty(key, value);
        }

        officeManager.execute(conversionTask);
    }

}
