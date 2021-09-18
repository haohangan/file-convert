package converter;

import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import org.jodconverter.core.document.DocumentFormat;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.core.util.AssertUtils;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

public class LocalFileConverter {

    private final LocalConverter converter;

    public LocalFileConverter(String openofficeHome, int portNumber) throws OfficeException {
        OfficeManager officeManager =
                LocalOfficeManager.builder()
                        .portNumbers(portNumber)
                        .officeHome(openofficeHome)
                        .build();
        officeManager.start();
        converter = LocalConverter.make(officeManager);
    }


    public void convert(InputStream is, OutputStream os, DocumentFormat format) throws OfficeException {
        converter
                .convert(is)
                .to(os).as(format)
                .execute();
    }


    public void convert(InputStream templateIs, Map<String, Object> paramMap, OutputStream os, DocumentFormat format) throws IOException, XDocReportException, OfficeException {
        AssertUtils.notNull(paramMap, "paramMap can not be null");
        IXDocReport report = XDocReportRegistry.getRegistry().loadReport(templateIs, TemplateEngineKind.Velocity);
        IContext context = report.createContext();
        paramMap.forEach(context::put);
        ByteArrayOutputStream arrayOs = new ByteArrayOutputStream();
        report.process(context, arrayOs);
        convert(new ByteArrayInputStream(arrayOs.toByteArray()), os, format);
    }


    public LocalConverter localConverter() {
        return converter;
    }

}
