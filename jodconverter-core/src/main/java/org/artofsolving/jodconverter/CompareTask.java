/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.artofsolving.jodconverter;

import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.Exception;
import java.io.File;
import java.util.HashMap;
import java.util.Map;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.OfficeContext;
import org.artofsolving.jodconverter.office.OfficeException;
import static org.artofsolving.jodconverter.office.OfficeUtils.SERVICE_DESKTOP;
import static org.artofsolving.jodconverter.office.OfficeUtils.cast;
import static org.artofsolving.jodconverter.office.OfficeUtils.toUnoProperties;
import static org.artofsolving.jodconverter.office.OfficeUtils.toUrl;

public class CompareTask extends AbstractOfficeTask {

    private final File compareTo;

    public CompareTask(File inputFile, File compareTo, File outputFile, DocumentFormat inputFormat, DocumentFormat outputFormat) {
        super(inputFile, outputFile, inputFormat, outputFormat);
        this.compareTo = compareTo;
    }

    @Override
    protected void modifyDocument(XComponent document, OfficeContext context) throws OfficeException {
        Object desktop = context.getService(SERVICE_DESKTOP);
        XMultiServiceFactory xFactory = context.getFactory();

        Object dispatchHelper;
        try {
            dispatchHelper = xFactory.createInstance("com.sun.star.frame.DispatchHelper");
        } catch (Exception ex) {
            throw new OfficeException("could not create instance", ex);
        }

        XDispatchHelper xDispatchHelper = cast(XDispatchHelper.class, dispatchHelper);

        XDesktop xDesktop = cast(XDesktop.class, desktop);
        XFrame xFrame = xDesktop.getCurrentFrame();

        XDispatchProvider xDispatchProvider = cast(XDispatchProvider.class, xFrame);

        Map<String, Object> compareProperties = new HashMap<String, Object>();
        compareProperties.put("URL", toUrl(compareTo));
        xDispatchHelper.executeDispatch(xDispatchProvider, ".uno:CompareDocuments", "", 0, toUnoProperties(compareProperties));
    }

}
