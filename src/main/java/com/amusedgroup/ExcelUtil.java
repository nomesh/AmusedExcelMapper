package com.amusedgroup;

import java.io.File;
import java.io.IOException;

public interface ExcelUtil {
    public File loadExcel(String filePath) throws IOException;

    public void replacePLayerNameAnomalies(File excelFile);
    public void findDuplicates(File excelFile);

}
