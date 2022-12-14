/*******************************************************************************
 * Copyright 2017 Bstek
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not
 * use this file except in compliance with the License.  You may obtain a copy
 * of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  See the
 * License for the specific language governing permissions and limitations under
 * the License.
 ******************************************************************************/
package com.ureport.ureportkeep.console.excel;


import com.opencsv.CSVReader;
import com.ureport.ureportkeep.console.AbstractReportBasicController;
import com.ureport.ureportkeep.console.cache.TempObjectCache;
import com.ureport.ureportkeep.console.exception.ReportDesignException;
import com.ureport.ureportkeep.console.html.HtmlPreviewController;
import com.ureport.ureportkeep.core.build.ReportBuilder;
import com.ureport.ureportkeep.core.definition.ReportDefinition;
import com.ureport.ureportkeep.core.exception.ReportComputeException;
import com.ureport.ureportkeep.core.exception.ReportException;
import com.ureport.ureportkeep.core.export.ExportConfigure;
import com.ureport.ureportkeep.core.export.ExportConfigureImpl;
import com.ureport.ureportkeep.core.export.ExportManager;
import com.ureport.ureportkeep.core.export.ReportRender;
import com.ureport.ureportkeep.core.export.excel.high.ExcelProducer;
import com.ureport.ureportkeep.core.model.Cell;
import com.ureport.ureportkeep.core.model.Column;
import com.ureport.ureportkeep.core.model.Report;
import com.ureport.ureportkeep.core.model.Row;
import com.ureport.ureportkeep.core.utils.ReportProperties;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.env.Environment;
import org.springframework.core.io.FileSystemResource;
import org.springframework.stereotype.Controller;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author Jacky.gao
 * @since 2017???4???17???
 */
@Controller
@RequestMapping(value = "/excel")
public class ExportExcelController extends AbstractReportBasicController {

    @Autowired
    private ReportRender reportRender;

    @Autowired
    private Environment env;

    @Autowired
    private ReportBuilder reportBuilder;

    @Autowired
    private ReportProperties reportProperties;

    @Autowired
    private ExportManager exportManager;

    private ExcelProducer excelProducer = new ExcelProducer();

    /**
     * ????????????excel
     * ????????????
     *
     * @param req
     * @param resp
     * @throws ServletException
     * @throws IOException
     */
    @RequestMapping(value = "/paging", method = RequestMethod.GET)
    public void paging(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        String dataxfilepath = env.getProperty("dataxfilepath");
        Object si = buildParameters(req).get("si");
        String file = req.getParameter("_u");

        //?????????????????????
        boolean ifSalary = req.getParameter("_u").contains("salary");
        StringBuilder titleParam = new StringBuilder("\\'");
        //???????????????????????????
        //????????????
        if (ifSalary) {
            //?????????????????????????????????dataX??????
            HashMap<String, Object> tmp_parameters = new HashMap<String, Object>() {{
                put("si", 1);
            }};
            ReportDefinition reportDefinition = reportRender.getReportDefinition(file);
            Report report = reportRender.render(reportDefinition, tmp_parameters);
            List<Row> rows = report.getRows();
            List<Column> columns = report.getColumns();
            Map<Row, Map<Column, Cell>> cellMap = report.getRowColCellMap();
            Map<Column, Cell> map = cellMap.get(rows.get(0));
            //??????????????????
            for (int i = 0; i < columns.size(); i++) {
                if (i == columns.size() - 1) {
                    titleParam.append(map.get(columns.get(i)).getFormatData().toString()).append("\\'");
                } else {
                    titleParam.append(map.get(columns.get(i)).getFormatData().toString()).append("\\',\\'");
                }
            }
        }
        String loadFileName = "";
        if (si == null && ifSalary && !"".equals(loadFileName = downloadDataxSalaryAll(titleParam))) {
            realDownloadAction(resp, dataxfilepath, loadFileName);
        } else if (ifSalary && Objects.requireNonNull(si).toString().contains(",") && !"".equals(loadFileName = downloadDataxSalaryMore(si.toString(), titleParam))) {
            //????????????????????????
            //System.out.println("download more salary data...");
            realDownloadAction(resp, dataxfilepath, loadFileName);
        } else {
            buildExcel(req, resp, true, false);
        }
    }


    /**
     * ??????dataX???????????????mysql???????????????,???????????????????????????
     *
     * @return
     */
    public String downloadDataxSalaryAll(StringBuilder title) {
        String dataxpy = env.getProperty("dataxpy");
        String json = env.getProperty("jsonpath") + env.getProperty("salary");
        String[] cmd = {"python", dataxpy, "-p -DtargetDir=" + env.getProperty("dataxfilepath") + " -Dtitle=" + title, json};
        return cmdPy(cmd);
    }


    /**
     * ?????????????????????????????????
     *
     * @return
     */
    public String downloadDataxSalaryMore(String params, StringBuilder title) {
        String dataxpy = env.getProperty("dataxpy");
        String json = env.getProperty("jsonpath") + env.getProperty("salary_more");
        String[] cmd = {"python", dataxpy, "-p -DtargetDir=" + env.getProperty("dataxfilepath") + " -Dstore_ids=" + params + " -Dtitle=" + title, json};
        return cmdPy(cmd);
    }

    public String cmdPy(String[] cmd) {
        String LoadedFileName = "";
        String line = "";
        try {
            Process process = Runtime.getRuntime().exec(cmd);
            BufferedReader in = new BufferedReader(new InputStreamReader(process.getInputStream()));
            while ((line = in.readLine()) != null) {
                if (line.contains("file name")) {
                    //??????dataX??????????????????
                    Pattern pattern = Pattern.compile("(?<=\\:\\[).*(?=\\])");
                    Matcher matcher = pattern.matcher(line);
                    if (matcher.find()) {
                        LoadedFileName = matcher.group();
                    }
                }
                System.out.println(line);
            }
            in.close();
            int re = process.waitFor();
            return re == 0 ? LoadedFileName : "";
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        }
    }

    public String getSingleFileName(String filename) {
        String dataxfilepath = env.getProperty("dataxfilepath");
        String readyFile = dataxfilepath + filename;
        File file = new File(readyFile);
        String okFile = readyFile + ".csv";
        file.renameTo(new File(okFile));
        return filename + ".csv";
    }

    public String CsvToElxs(String fileName) throws IOException {
        try {
            String dataxfilepath = env.getProperty("dataxfilepath");
            File file = new File(dataxfilepath + fileName);
            CSVReader reader = new CSVReader(new FileReader(file));
            SXSSFWorkbook workbook = new SXSSFWorkbook();
            BufferedOutputStream outputStream = new BufferedOutputStream(new FileOutputStream(dataxfilepath + fileName + "tmp"));
            SXSSFSheet sheet = workbook.createSheet();
            int index = 0;
            // ????????????
            for (String[] strings : reader) {
                SXSSFRow row = sheet.createRow(index++);
                for (int j = 0; j < strings.length; j++) {
                    row.createCell(j).setCellValue(strings[j]);
                }
            }
            workbook.write(outputStream);
            workbook.dispose();
            outputStream.close();
            reader.close();
            return fileName + "tmp";
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        }
    }

    /**
     * ????????????csv???????????????
     *
     * @return
     */
    public String getFileName() {
        try {
            File f = new File(Objects.requireNonNull(env.getProperty("dataxfilepath")));
            if (!f.exists()) {
                return null;
            }
            HashMap<String, Long> csvMap = new HashMap<>();
            File fa[] = f.listFiles();
            for (int i = 0; i < fa.length; i++) {//????????????
                File fs = fa[i];//?????????????????????i???
                if (!fs.isDirectory()) {
                    String fsName = fs.getName();
                    if (!fsName.contains("csv")) {
                        fs.renameTo(new File(fs + ".csv"));
                        csvMap.put(fs.getName() + ".csv", fs.lastModified());
                    } else {
                        csvMap.put(fs.getName(), fs.lastModified());
                    }
                }
            }
            //????????????????????????????????????????????????????????????????????????
            List<Map.Entry<String, Long>> list = new ArrayList(csvMap.entrySet());
            Collections.sort(list, (o1, o2) -> (int) (o1.getValue() - o2.getValue()));
            String maxFile = list.get(list.size() - 1).getKey();
            return CsvToElxs(maxFile);
        } catch (Exception e) {
            return "?????????????????????";
        }

    }

    /**
     * ?????????????????????
     */
    public void realDownloadAction(HttpServletResponse resp, String dataxfilepath, String loadFileName) throws IOException {
        //????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????
        String readyToLoad = "";
        if (!"".equals(loadFileName)) {
            readyToLoad = getSingleFileName(loadFileName);
        } else {
            readyToLoad = getFileName();
        }
        FileSystemResource file = new FileSystemResource(dataxfilepath + readyToLoad);
        String filename = file.getFilename();
        InputStream inputStream = null;
        BufferedInputStream bufferedInputStream = null;
        BufferedOutputStream bufferedOutputStream = null;
        resp.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(filename, "UTF-8"));
        try {
            inputStream = file.getInputStream();
            bufferedInputStream = new BufferedInputStream(inputStream);
            bufferedOutputStream = new BufferedOutputStream(resp.getOutputStream());
            FileCopyUtils.copy(bufferedInputStream, bufferedOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("<<the file download meet mistake>>");
        } finally {
            if (null != inputStream) {
                inputStream.close();
            }
            if (null != bufferedInputStream) {
                bufferedInputStream.close();
            }
            if (null != bufferedOutputStream) {
                bufferedOutputStream.flush();
                bufferedOutputStream.close();
            }
        }
    }

    /**
     * ???sheet??????excel
     *
     * @param req
     * @param resp
     * @throws ServletException
     * @throws IOException
     */
    @RequestMapping(value = {"/sheet", ""}, method = RequestMethod.GET)
    public void sheet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        buildExcel(req, resp, false, true);
    }

    public void buildExcel(HttpServletRequest req, HttpServletResponse resp, boolean withPage, boolean withSheet) throws IOException {
        String file = req.getParameter("_u");
        file = decode(file);
        if (StringUtils.isBlank(file)) {
            throw new ReportComputeException("Report file can not be null.");
        }
        OutputStream outputStream = resp.getOutputStream();
        try {
            String fileName = req.getParameter("_n");
            fileName = buildDownloadFileName(file, fileName, ".xlsx");
            resp.setContentType("application/octet-stream;charset=ISO8859-1");
            fileName = new String(fileName.getBytes("UTF-8"), "ISO8859-1");
            resp.setHeader("Content-Disposition", "attachment;filename=\"" + fileName + "\"");
            Map<String, Object> parameters = buildParameters(req);
            if (file.equals(PREVIEW_KEY)) {
                ReportDefinition reportDefinition = (ReportDefinition) TempObjectCache.getObject(PREVIEW_KEY);
                if (reportDefinition == null) {
                    throw new ReportDesignException("Report data has expired,can not do export excel.");
                }
                Report report = reportBuilder.buildReport(reportDefinition, parameters);
                if (withPage) {
                    excelProducer.produceWithPaging(report, outputStream);
                } else if (withSheet) {
                    excelProducer.produceWithSheet(report, outputStream);
                } else {
                    excelProducer.produce(report, outputStream);
                }
            } else {
                ExportConfigure configure = new ExportConfigureImpl(file, parameters, outputStream);
                if (withPage) {
                    //??????????????????
                    exportManager.exportExcelWithPaging(configure);
                } else if (withSheet) {
                    exportManager.exportExcelWithPagingSheet(configure);
                } else {
                    exportManager.exportExcel(configure);
                }
            }
        } catch (Exception ex) {
            throw new ReportException(ex);
        } finally {
            outputStream.flush();
            outputStream.close();
        }
    }

    public void setReportBuilder(ReportBuilder reportBuilder) {
        this.reportBuilder = reportBuilder;
    }

    public void setExportManager(ExportManager exportManager) {
        this.exportManager = exportManager;
    }

}
