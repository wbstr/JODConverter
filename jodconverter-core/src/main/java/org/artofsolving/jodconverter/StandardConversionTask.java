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
import org.artofsolving.jodconverter.document.DocumentFormat;

import com.sun.star.lang.XComponent;
import org.artofsolving.jodconverter.office.OfficeException;

public class StandardConversionTask extends AbstractOfficeTask {

    public StandardConversionTask(File inputFile, File outputFile, DocumentFormat inputFormat, DocumentFormat outputFormat) {
        super(inputFile, outputFile, inputFormat, outputFormat);
    }

    @Override
    protected void modifyDocument(XComponent document) throws OfficeException {
        // nop
    }

}
