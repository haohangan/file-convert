package org.example;

import fr.opensagres.xdocreport.converter.ConverterRegistry;
import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import fr.opensagres.xdocreport.converter.IConverter;
import fr.opensagres.xdocreport.converter.Options;
import fr.opensagres.xdocreport.converter.XDocConverterException;
import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.core.document.DocumentKind;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class OdtFile {

    public static void main(String[] args) throws IOException, XDocReportException {
// 1) Load ODT file and set Velocity template engine and cache it to the registry
        InputStream in = new FileInputStream("D:\\doc\\新建 OpenOffice 文字.odt");
        IXDocReport report = XDocReportRegistry.getRegistry().loadReport(in, TemplateEngineKind.Velocity);

// 2) Create Java model context
        IContext context = report.createContext();
        context.put("name", "编号0000009");

// 3) Generate report by merging Java model with the ODT
        OutputStream out = new FileOutputStream("D:\\doc\\新建 OpenOffice 文字2.odt");
        report.process(context, out);

        convert();

    }

    /**
     * xdocreport 的转pdf方案会存在一些格式上的误差
     */
    static void convert() throws FileNotFoundException, XDocConverterException {
        Options options = Options.getFrom(DocumentKind.ODT).to(ConverterTypeTo.PDF);

// 2) Get the converter from the registry
        IConverter converter = ConverterRegistry.getRegistry().getConverter(options);

// 3) Convert ODT 2 PDF
        InputStream in = new FileInputStream("D:\\doc\\新建 OpenOffice 文字2.odt");
        OutputStream out = new FileOutputStream("D:\\doc\\新建 OpenOffice 文字2.pdf");
        converter.convert(in, out, options);
    }
}
