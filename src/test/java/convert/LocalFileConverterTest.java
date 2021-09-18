package convert;

import converter.LocalFileConverter;
import fr.opensagres.xdocreport.core.XDocReportException;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.TimeUnit;

public class LocalFileConverterTest {

    static String path = "D:\\doc\\demo\\";

    public static void main(String[] args) throws OfficeException, IOException, XDocReportException {
        String home = "D:\\Program Files (x86)\\OpenOffice 4";
        int port = 8131;
        LocalFileConverter converter = new LocalFileConverter(home, port);
        InputStream is = transferStream(new FileInputStream(path + "新建 OpenOffice 文字.odt"));
        is.mark(Integer.MAX_VALUE);
        convert2Doc(converter, is);
        is.reset();
        convert2Pdf(converter, is);
        is.reset();
        convert2ReplacePdf(converter, is);
    }

    static void convert2Doc(LocalFileConverter converter, InputStream is) throws OfficeException {
        long start = System.nanoTime();
        converter.localConverter().convert(is, false).to(new File(path + "新建 OpenOffice 文字.doc")).execute();
        System.out.println(TimeUnit.SECONDS.convert(System.nanoTime() - start, TimeUnit.NANOSECONDS));
    }

    static void convert2Pdf(LocalFileConverter converter, InputStream is) {
        long start = System.nanoTime();
        converter.localConverter().convert(is, false).to(new File(path + "新建 OpenOffice 文字.pdf"));
        System.out.println(TimeUnit.SECONDS.convert(System.nanoTime() - start, TimeUnit.NANOSECONDS));
    }

    static void convert2ReplacePdf(LocalFileConverter converter, InputStream is) throws OfficeException, IOException, XDocReportException {
        long start = System.nanoTime();
        Map<String, Object> map = new HashMap<>();
        map.put("name", "罚现场000011");
        converter.convert(is, map, new FileOutputStream(path + "新建 OpenOffice 文字-replace.pdf"), DefaultDocumentFormatRegistry.PDF);
        System.out.println(TimeUnit.SECONDS.convert(System.nanoTime() - start, TimeUnit.NANOSECONDS));
    }


    static InputStream transferStream(InputStream stream) {
        if (stream.markSupported()) {
            return stream;
        }
        return new BufferedInputStream(stream);
    }
}
