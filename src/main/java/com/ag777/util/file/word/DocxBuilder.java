package com.ag777.util.file.word;

import com.ag777.util.file.FileUtils;
import com.ag777.util.file.word.config.TableItemConfig;
import com.ag777.util.file.word.config.WordStyleInterf;
import com.ag777.util.lang.IOUtils;
import com.ag777.util.lang.StringUtils;
import com.ag777.util.lang.exception.model.ImageNotSupportException;
import com.ag777.util.lang.function.TriConsumer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

/**
 * docx构建工具类(二次封装poi)
 * <p>
 * 依赖maven:
 * <ul>
 * <li>poi-ooxml</li>
 * <li>ooxml-schemas</li>
 * </ul>
 *
 * @author ag777
 * @version create on 2020年09月30日,last modify at 2023年05月15日
 */
public class DocxBuilder {

    private XWPFDocument doc;
    private final WordStyleInterf style;

    /** 代表目前刚创建一个新页(这个是为了解决newPage方法每次执行都会在新页插入个空行，我们在写入数据时可以复用该空行解决问题) */
    private boolean newPage;
    /** 第一级标题序号 */
    private int chapterNum;
    /** 第二级标题序号 */
    private int sectionNum;


    public DocxBuilder(WordStyleInterf wordStyleInterf) {
        this.style = wordStyleInterf;
        init();
    }

    public XWPFDocument doc() {
        return doc;
    }

    public void save(File file) throws IOException {
        FileOutputStream os = FileUtils.getOutputStream(file);
        try {
            doc.write(os);
        } finally {
            IOUtils.close(os, doc);
        }
    }

    private void init() {
        doc = DocxUtils.newDoc(
                style.fontFamily(),
                style.pageWidth(), style.pageHeight(),
                style.pageMarginTop(), style.pageMarginLeft(), style.pageMarginBottom(), style.pageMarginRight());
        chapterNum = 0;
        sectionNum = 0;
        newPage=true;
    }

    public DocxBuilder newPage() {
        DocxUtils.newPage(doc);
        newPage = true;
        return this;
    }

    public DocxBuilder title1(String title) {
        chapterNum++;
        sectionNum=0;
        if(chapterNum>1 && !newPage) {
            emptyLine(3, 12);
        }
        addTitle(chapterNum+"、"+title, 1, 11,true, false);
//        emptyLine(1, 12);
        return this;
    }

    public DocxBuilder title2(String title) {
        sectionNum++;
        addTitle(chapterNum+"."+sectionNum+"、"+title, 2, 10,true, true);
//        emptyLine(1, 12);
        return this;
    }

    public DocxBuilder text(String text, int fontSize, String rgbStr, boolean alignCenter, boolean isBold, boolean isItalic) {
        addText(text, fontSize, rgbStr, alignCenter, isBold, isItalic);
        return this;
    }

    public DocxBuilder emptyLine(int lineCount, int fontSize) {
        addEmptyLine(lineCount, fontSize);
        return this;
    }

    public DocxBuilder emptyLine(int lineCount) {
        addEmptyLine(lineCount);
        return this;
    }

    public <T> DocxBuilder table(List<T> dataList, List<TableItemConfig<T>> configList, boolean hasTitle) {
        addTable(dataList, configList, hasTitle);
        return this;
    }

    public <T> DocxBuilder vTable(Map<String, T> dataMap, int[] widths, TriConsumer<XWPFTableCell, T, String> contentRender, BiConsumer<XWPFTableCell, String> titleRender) {
        addVTable(dataMap, widths, contentRender, titleRender);
        return this;
    }

    public DocxBuilder img(File imgFile, Integer width, Integer height) {
        try {
            imgWithException(imgFile, width, height);
        } catch (Throwable ignored) {
        }
        return this;
    }

    public DocxBuilder imgWithException(File imgFile, Integer width, Integer height) throws InvalidFormatException, IOException, ImageNotSupportException {
        addImg(imgFile, width, height);
        return this;
    }

    /**
     * 添加页眉(默认居右)
     * @param text 页眉内容
     * @param fontSize 字体大小
     * @param rgbStr 字体颜色(16进制)
     * @return 页眉引用
     */
    public DocxBuilder header(String text, int fontSize, String rgbStr) {
        DocxUtils.header(doc,
                text,
                style.fontFamily(), fontSize, rgbStr);
        return this;
    }

    /**
     * 添加页脚(默认居中)
     * @param before 页数前的字符串
     * @param after 页数后的字符串
     * @param fontSize 字体大小
     * @param rgbStr 字体颜色(16进制)
     * @return 页脚引用
     */
    public DocxBuilder footer(String before, String after, int fontSize, String rgbStr) {
        DocxUtils.footer(doc,
                before, after,
                style.fontFamily(), fontSize, rgbStr);
        return this;
    }

    /**
     * 添加水印
     * @param text 水印文本
     * @param colorHex 16进制色值
     * @param width 水印宽度
     * @param height 水印高度
     * @param rotationAngle 旋转角度
     * @return DocxBuilder
     */
    public DocxBuilder waterMark(String text, String colorHex, int width, int height, int rotationAngle) {
       DocxUtils.addWaterMark(doc, text, colorHex, width, height, rotationAngle);
       return this;
    }

    public XWPFRun addText(String text, int fontSize, String rgbStr, boolean alignCenter, boolean isBold, boolean isItalic) {
        XWPFParagraph paragraph = newParagraph();
        return addText(paragraph, text, fontSize, rgbStr, alignCenter, isBold, isItalic);
    }

    public XWPFRun addText(XWPFParagraph paragraph, String text, int fontSize, String rgbStr, boolean alignCenter, boolean isBold, boolean isItalic) {
        XWPFRun run = DocxUtils.addText(paragraph, text, style.fontFamily(), fontSize, rgbStr);
        run.setBold(isBold);
        run.setItalic(isItalic);
        if(alignCenter) {
            paragraph.setAlignment(ParagraphAlignment.CENTER);
        }
        return run;
    }

    public XWPFRun addText(XWPFParagraph paragraph, String text, Integer size, String rgbStr) {
        return DocxUtils.addText(paragraph, text, style.fontFamily(), size, rgbStr, false);
    }

    public void addEmptyLine(int lineCount) {
        XWPFParagraph paragraph = doc.getLastParagraph();
        for (int i = 0; i < lineCount; i++) {
            paragraph.createRun().addBreak(BreakType.TEXT_WRAPPING);
        }
    }

    public void addEmptyLine(int lineCount, int fontSize) {
        XWPFParagraph paragraph = doc.createParagraph();
        String content="";
        if(lineCount > 1) {
            content= StringUtils.stack("\r", lineCount-1);
        }
        DocxUtils.addText(paragraph, content, style.fontFamily(), fontSize, null, true);
    }

    public void addImg(File imgFile, Integer width, Integer height) throws InvalidFormatException, IOException, ImageNotSupportException {
        XWPFParagraph paragraph = newParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        DocxUtils.addImg(paragraph, imgFile, width, height);
    }

    private XWPFParagraph newParagraph() {
        if(!newPage) {
            return doc.createParagraph();
        } else {
            newPage = false;
            return doc.getLastParagraph();
        }

    }

    /**
     * 创建横表
     * @param dataList 数据列表
     * @param configList 配置列表(列)
     * @param hasTitle 是否有标题栏,如果为false则不渲染标题栏
     * @param <T> 数据项类型
     * @return 表格引用
     */
    public <T>XWPFTable addTable(List<T> dataList, List<TableItemConfig<T>> configList, boolean hasTitle) {

        int[] widths = new int[configList.size()];
        for (int i = 0; i < configList.size(); i++) {
            widths[i] = configList.get(i).getWeight();
        }
        int rowNum = dataList.size();
        int colNum = configList.size();
        int baseRow = hasTitle?1:0;
        XWPFTable table = addTable(rowNum+baseRow, colNum, style.tableWidth(), widths);

        DocxUtils.fillTable(table, dataList, configList, style, hasTitle, (cell, config)->{
            /* 设置水平垂直居中 */
            DocxUtils.alignHCenter(cell);
            DocxUtils.alignVCenter(cell);
        });

        return table;
    }

    /**
     * 纵向表格
     * @param dataMap {"标题":"内容"}
     * @param widths [标题宽度,内容宽度]
     * @param contentRender 内容渲染器
     * @param titleRender 标题渲染器
     * @param <T> 内容类型
     * @return 表格对象
     */
    public <T>XWPFTable addVTable(Map<String, T> dataMap, int[] widths, TriConsumer<XWPFTableCell, T, String> contentRender, BiConsumer<XWPFTableCell, String> titleRender) {
        int titleCount = dataMap.keySet().size();
        XWPFTable table = addTable(titleCount, 2, style.tableWidth(), widths);
        Iterator<String> itor = dataMap.keySet().iterator();
        int i=0;
        while (itor.hasNext()) {
            String title = itor.next();
            T value = dataMap.get(title);
            XWPFTableRow row = table.getRow(i);
            XWPFTableCell cell = getCell(row, 0);
            titleRender.accept(cell, title);
            cell.setColor(style.colorTableTitleBg());

            cell = getCell(row, 1);
            contentRender.accept(cell, value, title);
            i++;
        }

        return table;
    }



    public XWPFTable addTable(int rowNum, int colNum, int tableWidth, int[] widths) {
        XWPFTable table = DocxUtils.newTable(doc, rowNum, colNum);
        DocxUtils.setBorder(table, XWPFTable.XWPFBorderType.SINGLE, 4, 0, "EFEFEF");
        DocxUtils.setTableWidths(table, tableWidth, widths);
        return table;
    }

    private XWPFRun addTitle(String title, int level, int fontSize, boolean isBold, boolean isItalic) {
        XWPFParagraph paragraph = newParagraph();
        paragraph.setStyle("Heading"+level);
        XWPFRun run = DocxUtils.addText(paragraph, title, style.fontFamily(), fontSize, null);
        run.setBold(isBold);
        run.setItalic(isItalic);
        run.addBreak(BreakType.TEXT_WRAPPING);
        return run;
    }

    private XWPFTableCell getCell(XWPFTableRow row, int colNum) {
        XWPFTableCell cell = row.getCell(colNum);
        /* 设置水平垂直居中 */
        DocxUtils.alignHCenter(cell);
        DocxUtils.alignVCenter(cell);
        return cell;
    }

    private static void setTableWidths(XWPFTable table, int[] colWidths) {
        CTTbl ttbl = table.getCTTbl();
        CTTblGrid tblGrid = ttbl.addNewTblGrid();

        for (Integer i : colWidths) {
            CTTblGridCol gridCol = tblGrid.addNewGridCol();

            gridCol.setW(BigInteger.valueOf(i));
        }
    }
}
