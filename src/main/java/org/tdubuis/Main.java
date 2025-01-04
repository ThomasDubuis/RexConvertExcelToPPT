package org.tdubuis;

import lombok.Getter;
import lombok.NonNull;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.TableCell;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.tdubuis.config.ConfigFile;
import org.tdubuis.filedata.ExcelData;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class Main {
    private static final Logger logger = LogManager.getLogger(Main.class);
    @Getter private static ConfigFile config;

    public static void main(String[] args) {
        if (args.length != 1) {
            logger.error("Need exactly 1 arguments <configFile>");
            return;
        }
        String configFileString = args[0];
        config = ConfigFile.parseConfigFile(new File(configFileString));

        if (config == null) {
            logger.error("Could not load config file: {}", configFileString);
            return;
        }

        String excelFileString = config.getExcelFile();
        String pptFileString = config.getPptFile();
        String outputFolderString = config.getOutputFolder();

        logger.info("Recap info : ");
        logger.info("Excel file : {}", excelFileString);
        logger.info("PPT file : {}", pptFileString);
        logger.info("Output folder : {}", outputFolderString);

        File excelFile = new File(excelFileString);
        File pptFile = new File(pptFileString);
        File outputFolder = new File(outputFolderString);

        if (!excelFile.exists() || !excelFile.isFile()) {
            logger.error("Excel file does not exist or is not a file");
            return;
        }
        if (!pptFile.exists() || !pptFile.isFile()) {
            logger.error("PPT file does not exist or is not a file");
            return;
        }
        if (!outputFolder.exists() || !outputFolder.isDirectory()) {
            logger.error("Output folder does not exist or is not a directory");
            return;
        }

        logger.info("Start Process");
        long startTime = System.currentTimeMillis();
        process(excelFile, pptFile, outputFolder);
        long endTime = System.currentTimeMillis();
        logger.info("Process completed in {} ms", endTime - startTime);
    }

    private static void process(File excelFile, File pptFile, File outputFolder) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
             XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(pptFile))) {

            logger.info("{} sheets found", workbook.getNumberOfSheets());
            logger.info("{} slides found", ppt.getSlides().size());

            Map<String, ExcelData> excelDataMap = new HashMap<>();
            for (int i = 0 ; i < workbook.getNumberOfSheets(); ++i) {
                addDataToExcelDataMap(workbook.getSheetAt(i), excelDataMap);
            }

            for (Map.Entry<String, ExcelData> entry : excelDataMap.entrySet()) {
                generatePPTWithExcelData(entry.getKey(), entry.getValue(), pptFile, outputFolder);
            }

            logger.info("End Process");
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }

    }

    private static void addDataToExcelDataMap(XSSFSheet sheetAt, Map<String, ExcelData> excelDataMap) {
        String sheetName = sheetAt.getSheetName();
        String[] sheetNameSplit = sheetName.split("-");
        String region = sheetNameSplit[0].trim();
        String sheetSlide = sheetNameSplit[1].trim();
        excelDataMap.putIfAbsent(region, new ExcelData(region));

        ExcelData excelData = excelDataMap.get(region);
        HashMap<String, ExcelData.Data> dataMap = new HashMap<>();
        addDataToExcelData(dataMap, sheetAt);

        if (sheetSlide.equals("MOIS")) {
            excelData.setDataMapMonth(dataMap);
        } else if (sheetSlide.equals("YTD")) {
            excelData.setDataMapYTD(dataMap);
        }else {
            throw new RuntimeException("Unsupported sheet slide: " + sheetSlide);
        }


    }

    private static void addDataToExcelData(HashMap<String, ExcelData.Data> dataMap, XSSFSheet sheet) {
        int numberOfRow = sheet.getLastRowNum();
        String currentConfigTitle = null;

        for (int i = 0; i < numberOfRow; i++) {
            XSSFRow row = sheet.getRow(i);
            String text = row.getCell(0).getStringCellValue();
            String configTitle = getConfig().isAndReturnConfigTitle(text);
            boolean endOfTable = text.trim().equalsIgnoreCase("RRF");

            if (configTitle != null) {
                currentConfigTitle = configTitle;
                dataMap.put(currentConfigTitle, new ExcelData.Data());
                continue;
            }

            if (currentConfigTitle != null) {
                addRowAndMergedRegionInExcelData(dataMap.get(currentConfigTitle), row, sheet);
            }else {
                logger.warn("No currentConfigTitle but not the end of file : {}", sheet.getSheetName());
                break;
            }

            if (endOfTable) {
                removeEmptyEndCell(dataMap.get(currentConfigTitle));
                currentConfigTitle = null;
            }
        }
    }
    private static void removeEmptyEndCell(ExcelData.Data data) {
        List<List<XSSFCell>> cellRows = data.getExcelCells();
        int lastIndexWithValue = 0;
        for (List<XSSFCell> row : cellRows) {
            for (int i = row.size() - 1; i >= 0; i--) {
                if (!row.get(i).toString().isBlank()) {
                    if (lastIndexWithValue < i) {
                        lastIndexWithValue = i;
                    }
                    break;
                }
            }
        }
        for (List<XSSFCell> row : cellRows) {
            row.removeAll(row.subList(lastIndexWithValue + 1, row.size()));
        }

    }

    private static void addRowAndMergedRegionInExcelData(ExcelData.Data data, XSSFRow row, XSSFSheet sheet) {
        List<XSSFCell> cellRow = new ArrayList<>();
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (int i = 0; i < row.getLastCellNum(); i++) {
            XSSFCell cell = row.getCell(i);
            cellRow.add(cell);

            CellRangeAddress mergedRegion = getMergedRegionIfIsMergedCell(cell, mergedRegions);
            if (mergedRegion != null && !data.getMergedRegion().contains(mergedRegion)) {
                data.getMergedRegion().add(mergedRegion);
            }
            data.getColumnWidth().putIfAbsent(i, sheet.getColumnWidthInPixels(i));
        }
        data.getExcelCells().add(cellRow);
    }

    private static CellRangeAddress getMergedRegionIfIsMergedCell(XSSFCell cell, List<CellRangeAddress> mergedRegions) {
        for (CellRangeAddress mergedRegion : mergedRegions) {
            if (mergedRegion.isInRange(cell)) {
                return mergedRegion;
            }
        }
        return null;
    }

    private static void generatePPTWithExcelData(String pptName, ExcelData excelData, File pptFile, File outputFolder) {
        logger.debug("Generate PPT : {}", pptName);
        try (XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(pptFile))) {

            for (ConfigFile.Config config : getConfig().getConfig()) {
                ConfigFile.Position position = config.getPosition();
                Integer textSize = config.getTextSize();
                ExcelData.Data dataMonth = excelData.getDataMapMonth().get(config.getTitle());
                ExcelData.Data dataYTD = excelData.getDataMapYTD().get(config.getTitle());

                if (dataMonth == null) {
                    logger.error("Error when generate slide month {}, abort this region {}", config.getSlideMonth(), pptName);
                } else {
                    generateSlide(dataMonth, ppt, position, textSize, config.getSlideMonth()); //Generate slide Month
                }
                if (dataYTD == null) {
                    logger.error("Error when generate slide YTD {} for this region : {}", config.getSlideYTD(), pptName);
                } else {
                    generateSlide(dataYTD, ppt, position, textSize, config.getSlideYTD()); //Generate slide YTD
                }
            }

            FileOutputStream out = new FileOutputStream(outputFolder.getAbsolutePath() + "/" + pptName + getConfig().getExcelSuffix() + ".pptx");
            ppt.write(out);
            out.close();
            throw new RuntimeException("Fin for not generating all Files"); //TODO Remove this thing
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static void generateSlide(@NonNull ExcelData.Data data, @NonNull XMLSlideShow ppt, @NonNull ConfigFile.Position position, @NonNull Integer textSize, @NonNull Integer slidePos) {
        logger.debug("Generate slide : {}", slidePos);

        XSLFSlide slide = ppt.getSlides().get(slidePos - 1);
        XSLFTable table = slide.createTable();
        table.setAnchor(new Rectangle(position.getX(), position.getY(), position.getWidth(), position.getHeight()));

        //Add Data and Style
        for (List<XSSFCell> excelRow : data.getExcelCells()) {
            XSLFTableRow pptRow = table.addRow();
            for (XSSFCell excelCell : excelRow) {
                XSLFTableCell pptCell = pptRow.addCell();
                copyExcelCellToPptCell(pptCell, excelCell, pptRow);
            }
        }
        //Set Column width
        for (Map.Entry<Integer, Float> entry : data.getColumnWidth().entrySet()) {
            if (entry.getKey() >= table.getNumberOfColumns()) {
                break;
            }
            table.setColumnWidth(entry.getKey(), entry.getValue());
        }

        //MergeRegion
        data.shiftMergedRegionToOrigin();
        for (CellRangeAddress mergedRegion : data.getMergedRegion()) {
            table.mergeCells(mergedRegion.getFirstRow(), mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
        }

        //Fix bug Border not working when cell are merged
        //https://bz.apache.org/bugzilla/show_bug.cgi?id=62431
        fixBorderOnMergedCell(table);
    }

    private static void fixBorderOnMergedCell(XSLFTable table) {
        for (int rowPos = 0; rowPos < table.getRows().size() ; rowPos++) {
            XSLFTableRow pptRow = table.getRows().get(rowPos);
            for (int colPos = 0; colPos < pptRow.getCells().size(); colPos++) {
                XSLFTableCell cell = pptRow.getCells().get(colPos);
                if (cell.isMerged()) {
                    for (TableCell.BorderEdge borderEdge : TableCell.BorderEdge.values()) {
                        if (cell.getBorderWidth(borderEdge) != null) {
                            switch (borderEdge) {
                                case top -> copyStyleToCloseCell(borderEdge, cell, table.getCell(rowPos - 1, colPos), TableCell.BorderEdge.bottom);
                                case bottom -> copyStyleToCloseCell(borderEdge, cell, table.getCell(rowPos + 1, colPos), TableCell.BorderEdge.top);
                                case left -> copyStyleToCloseCell(borderEdge, cell, table.getCell(rowPos, colPos - 1), TableCell.BorderEdge.right);
                                case right -> copyStyleToCloseCell(borderEdge, cell, table.getCell(rowPos, colPos + 1), TableCell.BorderEdge.left);
                            }
                        }
                    }
                }
            }
        }
    }

    private static void copyStyleToCloseCell(TableCell.BorderEdge borderEdge, XSLFTableCell cell, XSLFTableCell cellNextTo, TableCell.BorderEdge oppositeBorderEdge) {
        if (cellNextTo == null) return;
        cellNextTo.removeBorder(oppositeBorderEdge);
        cellNextTo.setBorderWidth(oppositeBorderEdge, cell.getBorderWidth(borderEdge));
        cellNextTo.setBorderColor(oppositeBorderEdge, cell.getBorderColor(borderEdge));
        cell.removeBorder(borderEdge);
    }


    private static void copyExcelCellToPptCell(XSLFTableCell pptCell, XSSFCell excelCell, XSLFTableRow pptRow) {
        XSSFCellStyle excelStyle = excelCell.getCellStyle();
        //BackgroundColor
        if (excelStyle.getFillForegroundColorColor() != null) {
            int[] rgb = convertRGBByteToRGBInt(excelStyle.getFillForegroundColorColor().getRGB());
            Color color = new Color(rgb[0], rgb[1], rgb[2]);
            pptCell.setFillColor(color);
        }
        //VerticalAlignment
        pptCell.setVerticalAlignment(convertVerticalAlignment(excelStyle.getVerticalAlignment()));

        //Text
        XSLFTextParagraph textParagraph = pptCell.addNewTextParagraph();
        textParagraph.setTextAlign(convertHorizontalAlignment(excelStyle.getAlignment()));
        XSLFTextRun textRun = textParagraph.addNewTextRun();
        textRun.setText(formatText(excelCell.toString(), excelStyle.getDataFormatString()));
        Font font = excelStyle.getFont();
        textRun.setFontFamily(font.getFontName());
        if (excelStyle.getFont().getCTFont().getColorArray(0).getRgb() != null) {
            int[] fontRgb = convertRGBByteToRGBInt(excelStyle.getFont().getCTFont().getColorArray(0).getRgb());
            textRun.setFontColor(new Color(fontRgb[1],fontRgb[2], fontRgb[3], fontRgb[0]));
        }
        textRun.setFontSize(9d);
        textRun.setBold(font.getBold());
        textRun.setItalic(font.getItalic());

        //Border
        applyBorderStyle(pptCell, excelStyle);

        //Fix Bug borderLeft not working
        //https://bz.apache.org/bugzilla/show_bug.cgi?id=69501
        fixLeftBorderStyle(pptCell, excelStyle, pptRow);
    }

    private static void fixLeftBorderStyle(XSLFTableCell pptCell, XSSFCellStyle excelStyle, XSLFTableRow pptRow) {
        BorderStyle borderLeftStyle = excelStyle.getBorderLeft();
        if (borderLeftStyle != BorderStyle.NONE) {
            for (int i = 0; i < pptRow.getCells().size(); i++) {
                if (pptRow.getCells().get(i).equals(pptCell) && i > 0) {
                    int[] rgb = convertRGBByteToRGBInt(excelStyle.getLeftBorderXSSFColor().getRGB());
                    XSLFTableCell pptRowPrevious = pptRow.getCells().get(i-1);
                    pptRowPrevious.removeBorder(TableCell.BorderEdge.right);
                    pptRowPrevious.setBorderColor(TableCell.BorderEdge.right, new Color(rgb[0],rgb[1],rgb[2]));
                    pptRowPrevious.setBorderWidth(TableCell.BorderEdge.right, 1);
                    return;
                }
            }
        }
    }

    private static String formatText(String text, String format) {
        if(text.isBlank()) {
            return text;
        }
        return switch (format) {
            case "mmm-yy" -> formatDate(text);
            case "0%", "#,##0" -> formatNumber(text, format);
            case "General" -> formatGeneral(text);
            default -> throw new IllegalStateException("Unexpected format value : " + format);
        };
    }
    private static String formatGeneral(String text) {
        try {
            double number = Double.parseDouble(text);
            if (number == (int) number) {
                return String.valueOf((int) number);
            } else {
                return String.valueOf(number);
            }
        } catch (NumberFormatException e) {
            return text;
        }

    }

    private static String formatNumber(String text, String format) {
        return new DecimalFormat(format).format(Float.valueOf(text));
    }

    private static String formatDate(String text) {
        try {
            SimpleDateFormat inputFormat = new SimpleDateFormat("dd-MMMM-yyyy");
            Date date = inputFormat.parse(text);
            SimpleDateFormat outputFormat = new SimpleDateFormat("MMM-yy");
            return outputFormat.format(date);
        } catch (Exception e) {
            logger.error(e);
            throw new IllegalStateException("FormatDate has failed text : " + text);
        }
    }

    public static int[] convertRGBByteToRGBInt(byte[] b) {
        if (b == null || (b.length != 3 && b.length != 4)) {
            logger.error("Need exactly 3 or 4 bytes to get RGB int");
            return new int[0];
        }
        int[] rgb = new int[b.length];
        for (int i = 0; i < b.length; i++) {
            rgb[i] = b[i] & 0xFF;
        }
        return rgb;
    }

    private static void applyBorderStyle(XSLFTableCell pptCell, XSSFCellStyle excelStyle) {
        //TODO les border ne fonctionnent pas correctement

        BorderStyle borderBottomStyle = excelStyle.getBorderBottom();
        if (borderBottomStyle != BorderStyle.NONE) {
            int[] rgb = convertRGBByteToRGBInt(excelStyle.getBottomBorderXSSFColor().getRGB());
            pptCell.setBorderColor(TableCell.BorderEdge.bottom, new Color(rgb[0],rgb[1],rgb[2]));
            pptCell.setBorderWidth(TableCell.BorderEdge.bottom, 1);
        }

        BorderStyle borderTopStyle = excelStyle.getBorderTop();
        if (borderTopStyle != BorderStyle.NONE) {
            int[] rgb = convertRGBByteToRGBInt(excelStyle.getTopBorderXSSFColor().getRGB());
            pptCell.setBorderColor(TableCell.BorderEdge.top, new Color(rgb[0],rgb[1],rgb[2]));
            pptCell.setBorderWidth(TableCell.BorderEdge.top, 1);
        }
//      Note : The setBorderLeft not working (ApachePOI Bug) so I fix with other way
//        BorderStyle borderLeftStyle = excelStyle.getBorderLeft();
//        if (borderLeftStyle != BorderStyle.NONE) {
//            int[] rgb = convertRGBByteToRGBInt(excelStyle.getLeftBorderXSSFColor().getRGB());
//            pptCell.setBorderColor(TableCell.BorderEdge.left, new Color(rgb[0],rgb[1],rgb[2]));
//            pptCell.setBorderWidth(TableCell.BorderEdge.left, 1);
//        }

        BorderStyle borderRightStyle = excelStyle.getBorderRight();
        if (borderRightStyle != BorderStyle.NONE) {
            int[] rgb = convertRGBByteToRGBInt(excelStyle.getRightBorderXSSFColor().getRGB());
            pptCell.setBorderColor(TableCell.BorderEdge.right, new Color(rgb[0],rgb[1],rgb[2]));
            pptCell.setBorderWidth(TableCell.BorderEdge.right, 1);
        }
    }

    private static VerticalAlignment convertVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment verticalAlignment) {
        return switch (verticalAlignment) {
            case CENTER -> VerticalAlignment.MIDDLE;
            case TOP -> VerticalAlignment.TOP;
            case BOTTOM -> VerticalAlignment.BOTTOM;
            case JUSTIFY -> VerticalAlignment.JUSTIFIED;
            case DISTRIBUTED -> VerticalAlignment.DISTRIBUTED;
        };
    }

    private static TextParagraph.TextAlign convertHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        return switch (horizontalAlignment) {
           case CENTER, CENTER_SELECTION -> TextParagraph.TextAlign.CENTER;
            case DISTRIBUTED -> TextParagraph.TextAlign.DIST;
            case JUSTIFY -> TextParagraph.TextAlign.JUSTIFY;
            case LEFT -> TextParagraph.TextAlign.LEFT;
            case RIGHT -> TextParagraph.TextAlign.RIGHT;
            case FILL -> TextParagraph.TextAlign.THAI_DIST;
            case GENERAL -> TextParagraph.TextAlign.JUSTIFY_LOW;
        };
    }
}