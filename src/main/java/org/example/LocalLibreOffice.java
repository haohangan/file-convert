package org.example;

import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.core.office.OfficeUtils;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;

import java.io.File;

public class LocalLibreOffice {

    public static void main(String[] args) throws OfficeException {
        OfficeManager officeManager =
                LocalOfficeManager.builder()
                        .portNumbers(8013)
                        .officeHome("D:\\Program Files (x86)\\OpenOffice 4")
                        .build();
        try {
            officeManager.start();
            LocalConverter converter = LocalConverter.make(officeManager);
            converter
                    .convert(new File("D:\\doc\\新建 OpenOffice 文字2.odt"))
                    .to(new File("D:\\doc\\新建 OpenOffice 文-openoffice.pdf"))
                    .execute();
        } finally {
            // Stop the office process
            OfficeUtils.stopQuietly(officeManager);
        }
    }
}
