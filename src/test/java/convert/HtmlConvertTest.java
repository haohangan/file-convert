package convert;

import converter.LocalFileConverter;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.TimeUnit;

public class HtmlConvertTest {
    static String path = "D:\\jsworkspace\\html\\";
    static String content = "<!DOCTYPE html>\n" +
            "<html>\n" +
            "\n" +
            "<head>\n" +
            "    <meta charset=\"utf-8\">\n" +
            "    <meta http-equiv=\"Content-Type\" content=\"text/html;charset=UTF-8\">\n" +
            "    <title>2022年5月中国采购经理指数运行情况</title>\n" +
            "</head>\n" +
            "\n" +
            "<body>\n" +
            "    <h2 class=\"xilan_tit\">2022年5月中国采购经理指数运行情况</h2>\n" +
            "    <p class=\"MsoNormal\"><span style=\"font-size:\n" +
            "                12pt; font-family:\n" +
            "                宋体\"> 供应商配送时间指数为</span><span lang=\"EN-US\" style=\"font-size: 12pt;\n" +
            "                font-family: &quot;Times New\n" +
            "                Roman&quot;,&quot;serif&quot;\">44.1%</span><span style=\"font-size: 12pt; font-family:\n" +
            "                宋体\">，比上月上升</span><span lang=\"EN-US\" style=\"font-size: 12pt;\n" +
            "                font-family:\n" +
            "                &quot;Times New\n" +
            "                Roman&quot;,&quot;serif&quot;\">6.9</span><span style=\"font-size:\n" +
            "                12pt; font-family:\n" +
            "                宋体\">个百分点，仍低于临界点，表明制造业原材料供应商交货时间仍然较慢。</span></p>\n" +
            "</body>";

    public static void main(String[] args) throws OfficeException, IOException {
        String home = "D:\\Program Files (x86)\\OpenOffice 4";
        int port = 8131;
        LocalFileConverter converter = new LocalFileConverter(home, port);
//        convert2Doc(converter, path + "a.html");
//        convert2Pdf(converter, path + "a.html");

        convert2Doc(converter, new ByteArrayInputStream(content.getBytes(StandardCharsets.UTF_8)));
    }

    static void convert2Doc(LocalFileConverter converter, String filePath) throws OfficeException {
        long start = System.nanoTime();
        converter.localConverter().convert(new File(filePath)).as(DefaultDocumentFormatRegistry.XHTML).to(new File(path + "新建 OpenOffice 文字.doc")).as(DefaultDocumentFormatRegistry.DOC).execute();
        System.out.println(TimeUnit.SECONDS.convert(System.nanoTime() - start, TimeUnit.NANOSECONDS));
    }

    static void convert2Pdf(LocalFileConverter converter, InputStream is) throws OfficeException {
        long start = System.nanoTime();
        converter.localConverter().convert(is).as(DefaultDocumentFormatRegistry.XHTML).to(new File(path + "新建 OpenOffice 文字.pdf")).as(DefaultDocumentFormatRegistry.PDF).execute();
        System.out.println(TimeUnit.SECONDS.convert(System.nanoTime() - start, TimeUnit.NANOSECONDS));
    }


    static void convert2Doc(LocalFileConverter converter, InputStream is) throws OfficeException {
        long start = System.nanoTime();
        converter.localConverter().convert(is).as(DefaultDocumentFormatRegistry.HTML).to(new File(path + "新建 OpenOffice 文字.doc")).as(DefaultDocumentFormatRegistry.DOC).execute();
        System.out.println(TimeUnit.SECONDS.convert(System.nanoTime() - start, TimeUnit.NANOSECONDS));
    }

    static void convert2Pdf(LocalFileConverter converter, String filePath) throws OfficeException {
        long start = System.nanoTime();
        converter.localConverter().convert(new File(filePath)).as(DefaultDocumentFormatRegistry.HTML).to(new File(path + "新建 OpenOffice 文字.pdf")).as(DefaultDocumentFormatRegistry.PDF).execute();
        System.out.println(TimeUnit.SECONDS.convert(System.nanoTime() - start, TimeUnit.NANOSECONDS));
    }

}
