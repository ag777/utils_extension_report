package com.ag777.util.file.word;

import com.ag777.util.file.word.config.TableItemConfig;
import com.ag777.util.file.word.config.WordStyleInterf;
import com.ag777.util.lang.exception.model.ImageNotSupportException;
import com.ag777.util.lang.function.TriConsumer;
import com.ag777.util.lang.img.ImageUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 有关word文档操作工具类(二次封装poi)
 * <p>
 * 依赖maven:
 * <ul>
 * <li>poi-ooxml</li>
 * <li>ooxml-schemas</li>
 * </ul>
 *
 * @author ag777
 * @version create on 2020年09月30日,last modify at 2020年09月30日
 */
public class DocxUtils {

//    public static final String FONT_FAMILY_DEFAULT = "微 软 雅 黑";

    private static final Pattern P_ENTER = Pattern.compile("\r?\n");    //前面可以有\r的\n

    /**
     * cell水平居中
     * @param cell 单元格
     */
    public static void alignHCenter(XWPFTableCell cell) {
        alignH(cell, STJc.CENTER);
    }

    /**
     * cell垂直居中
     * @param cell 单元格
     */
    public static void alignVCenter(XWPFTableCell cell) {
        alignV(cell, STVerticalJc.CENTER);
    }

    /**
     * cell水平状态
     * @param cell
     * @param stjc
     */
    public static void alignH(XWPFTableCell cell, STJc.Enum stjc) {
        CTTc cttc = cell.getCTTc();
        CTP ctp = cttc.getPList().get(0);
        CTPPr ctppr = ctp.isSetPPr()?ctp.getPPr():ctp.addNewPPr();
        CTJc ctjc = ctppr.isSetJc()?ctppr.getJc():ctppr.addNewJc();

        ctjc.setVal(stjc); //水平居中
    }

    /**
     * cell垂直状态
     * @param cell
     * @param stverticalJc
     */
    public static void alignV(XWPFTableCell cell, STVerticalJc.Enum stverticalJc) {
        CTTcPr ctPr = getPr(cell);
        ctPr.addNewVAlign().setVal(stverticalJc);
    }


    /**
     *
     * @param doc doc
     * @return 获取doc的属性
     */
    public static CTSectPr getPr(XWPFDocument doc) {
        CTBody body = doc.getDocument().getBody();
        return body.isSetSectPr()? body.getSectPr():body.addNewSectPr();
    }

    /**
     *
     * @param table table
     * @return 获取table属性
     */
    public static CTTblPr getPr(XWPFTable table) {
        CTTbl ttbl = table.getCTTbl();
        return ttbl.getTblPr() != null ? ttbl.getTblPr(): ttbl.addNewTblPr();
    }

    /**
     *
     * @param cell cell
     * @return 获取cell属性
     */
    public static CTTcPr getPr(XWPFTableCell cell) {
        CTTc cttc = cell.getCTTc();
        return cttc.isSetTcPr()?cttc.getTcPr(): cttc.addNewTcPr();
    }

    /**
     *
     * @param run XWPFRun
     * @return
     */
    public static CTRPr getPr(XWPFRun run) {
        CTR ctr = run.getCTR();
        return ctr.isSetRPr()?ctr.getRPr():ctr.addNewRPr();
    }

    public static XWPFParagraph getParagraph(XWPFTableCell cell) {
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        return paragraphs.get(0);
    }

//    /**
//     * 设置字体大小
//     * @param pr CTRPr
//     * @param fontSize 字体大小
//     */
//    public static void setFontSize(CTRPr pr, long fontSize) {
//        CTHpsMeasure sz = pr.isSetSz() ? pr.getSz() : pr.addNewSz();
//        sz.setVal(BigInteger.valueOf(fontSize));
//        sz = pr.isSetSzCs() ? pr.getSzCs() : pr
//                .addNewSzCs();
//        sz.setVal(BigInteger.valueOf(fontSize));
//    }

    public static XWPFDocument newDoc(
            String fontFamily,
            long pageWidth, long pageHeight,
            long marginTop, long marginLeft, long marginBottom, long marginRight) {
        XWPFDocument doc = new XWPFDocument();
        CTSectPr section = DocxUtils.getPr(doc);
        if(!section.isSetPgSz()) {
            section.addNewPgSz();
        }
        // 设置页面大小  当前A4大小
        CTPageSz pageSize = section.getPgSz();
        pageSize.setW(BigInteger.valueOf(pageWidth));
        pageSize.setH(BigInteger.valueOf(pageHeight));
        pageSize.setOrient(STPageOrientation.PORTRAIT); //竖版,横版是LANDSCAPE

        CTPageMar pageMar = section.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(marginLeft));
        pageMar.setTop(BigInteger.valueOf(marginTop));
        pageMar.setRight(BigInteger.valueOf(marginRight));
        pageMar.setBottom(BigInteger.valueOf(marginBottom));

        //设置全局字体
        XWPFStyles styles = doc.createStyles();
        CTFonts fonts = CTFonts.Factory.newInstance();
        fonts.setEastAsia(fontFamily);
        fonts.setHAnsi(fontFamily);

        styles.setDefaultFonts(fonts);
        return doc;
    }

    /**
     * 加入分页符
     * @param doc doc
     */
    public static void newPage(XWPFDocument doc) {
        //这种会在新页顶端多一个空行,https://blog.csdn.net/john1337/article/details/104900715
//        XWPFParagraph p = doc.createParagraph();
        //实验证明两种都会有空行,并且这里不能用getLastParagraph,因为(不新建段落的话)样式会互相影响
        doc.createParagraph().createRun().addBreak(BreakType.PAGE);
    }

    public static XWPFPicture addImg(XWPFDocument doc, File imgFile, Integer width, Integer height) throws IOException, InvalidFormatException, ImageNotSupportException {
        XWPFParagraph paragraph = doc.createParagraph();
        return addImg(paragraph, imgFile, width, height);
    }

    /**
     * 单元格中插入图片
     * @param paragraph cell
     * @param imgFile 图片文件
     * @param width 宽度
     * @param height 高度
     * @throws IOException
     * @throws InvalidFormatException
     * @throws ImageNotSupportException
     * @return XWPFRun
     */
    public static XWPFPicture addImg(XWPFParagraph paragraph, File imgFile, Integer width, Integer height) throws IOException, InvalidFormatException, ImageNotSupportException {
        XWPFRun run = paragraph.createRun();
        return addImg(run, imgFile, width, height);
    }

    public static XWPFPicture addImg(XWPFRun run, File imgFile, Integer width, Integer height) throws IOException, InvalidFormatException, ImageNotSupportException {
        if(!imgFile.exists()){
            return null;
        }
        int format = getFormat(imgFile, XWPFDocument.PICTURE_TYPE_PNG);
        if(width == null || height == null) {
            int[] groups = ImageUtils.getWidthAndHeight(imgFile.getAbsolutePath());
            if(width == null && height == null) {
                width = groups[0];
                height = groups[1];

                width = Math.round(width*0.8f);
                height = Math.round(height*0.8f);

            } else if(width == null) {  //(w/w原)=(h/h原) w=(w原*h/h原)
                width = Math.round(groups[0] * (height*1f/groups[1]));
            } else {    //height == null
                height = Math.round(groups[1] * (width*1f/groups[0]));
            }

        }

        try (FileInputStream is = new FileInputStream(imgFile)) {
            return run.addPicture(is, format, imgFile.getName(), Units.toEMU(width), Units.toEMU(height)); // 200x200 pixels
        }
    }

    /**
     * 插入页码
     * @param paragraph 段落
     * @param fontSize 字体大小
     * @param rgbStr 字体颜色
     */
    public static void addPageNum(XWPFParagraph paragraph, String fontFamily, int fontSize, String rgbStr) {
        XWPFRun run = addText(paragraph, "", fontFamily, fontSize, rgbStr); //这句是为了填充样式
        CTR ctr = run.getCTR();
        //start
        CTFldChar fldChar = ctr.addNewFldChar();
        fldChar.setFldCharType(STFldCharType.Enum.forString("begin"));
        //页码
        CTText ctText = ctr.addNewInstrText();
        ctText.setStringValue("PAGE  \\* MERGEFORMAT");
        ctText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
        //end
        fldChar = ctr.addNewFldChar();
        fldChar.setFldCharType(STFldCharType.Enum.forString("end"));
    }

    public static XWPFRun addText(XWPFTableCell cell, String text, String fontFamily, Integer size, String rgbStr) {
        XWPFParagraph paragraph = getParagraph(cell);
        return addText(paragraph, text, fontFamily, size, rgbStr, false);
    }
    
    public static XWPFRun addText(XWPFParagraph paragraph, String text, String fontFamily, Integer size, String rgbStr) {
        return addText(paragraph, text, fontFamily, size, rgbStr, false);
    }

    public static XWPFRun addText(XWPFParagraph paragraph, String text, String fontFamily, Integer size, String rgbStr, boolean isBold) {
        XWPFRun run = paragraph.createRun();
        if(text != null) {
            text = P_ENTER.matcher(text).replaceAll("\r");  //\r换行, \r\n会有缩进,只有\n不换行,这里的\n类似于\t的效果
        } else {
            text = "";
        }

        run.setText(text);
        run.setFontFamily(fontFamily);

        if(size != null) {  //字体
            run.setFontSize(size);
        }

        if(rgbStr != null) {    //颜色
            run.setColor(rgbStr);
        }

        if(isBold) {    //粗体
            run.setBold(isBold);
        }

        return run;
    }

//    public static XWPFRun addParagraph(XWPFParagraph paragraph, String text, String fontFamily, Integer size, String rgbStr, boolean isBold) {
//        XWPFRun run = paragraph.createRun();
//        if(text != null) {
//
//        } else {
//            text = "";
//        }
//        run.setText(text);
//        run.setFontFamily(fontFamily);
//
//        if(size != null) {  //字体
//            run.setFontSize(size);
//        }
//
//        if(rgbStr != null) {    //颜色
//            run.setColor(rgbStr);
//        }
//
//        if(isBold) {    //粗体
//            run.setBold(isBold);
//        }
//
//        return run;
//    }


    public static <T>XWPFTable fillTable(XWPFTable table, List<T> dataList, List<TableItemConfig<T>> configList, WordStyleInterf style, boolean hasTitle, BiConsumer<XWPFTableCell, TableItemConfig<T>> onCreateCell) {
        int colNum = configList.size();
        int baseRow = hasTitle?1:0;

        /*列表标题栏*/
        if(hasTitle) {
            XWPFTableRow row = table.getRow(0);
            DocxUtils.setRepeatHeaderRow(row);

            for (int i = 0; i < colNum; i++) {
                XWPFTableCell cell = row.getCell(i);
                cell.setColor(style.colorTableTitleBg());

                //配置项中获取标题和渲染器
                TableItemConfig<T> config = configList.get(i);
                if(onCreateCell != null) {
                    onCreateCell.accept(cell, config);
                }
                BiConsumer<XWPFTableCell, String> render = config.getTitleRender();
                if(render != null) {
                    render.accept(cell, config.getTitle());
                } else {    //没有设定的情况下，用默认样式渲染
                    XWPFParagraph paragraph = DocxUtils.getParagraph(cell);
                    DocxUtils.addText(
                            paragraph,
                            config.getTitle(),
                            style.fontFamily(),
                            style.tableTitleFontSize(), null);
                }
            }
        }

        /*列表主体*/
        for (int i=0;i < dataList.size(); i++) {
            XWPFTableRow row = table.getRow(i+baseRow);
            T item = dataList.get(i);
            for (int j = 0; j < colNum; j++) {
                XWPFTableCell cell = row.getCell(j);
                //配置项中获取标题和渲染器
                TableItemConfig<T> config = configList.get(j);
                if(onCreateCell != null) {
                    onCreateCell.accept(cell, config);
                }
                if(config.getAlign() != STJc.CENTER) {  //默认居中
                    DocxUtils.alignH(cell, config.getAlign());
                }
                BiFunction<T, String, String> getContent = config.getGetContent();
                if(getContent != null) {    //直接取内容展示
                    String content = getContent.apply(item, config.getKey());
                    DocxUtils.addText(
                            cell,
                            content,
                            style.fontFamily(),
                            style.tableContentFontSize(), config.getColorHex());
                } else {    //自定义内容展示
                    TriConsumer<XWPFTableCell, T, String> render = config.getContentRender();
                    render.accept(cell, item, config.getKey());
                }
            }
        }

        return table;
    }

    /**
     * 构造一个rowNum行colNum列的表格，注意，这时表格是没有宽度的
     * 默认cellMargin为 上下左右100
     * @param doc doc
     * @param rowNum 行数
     * @param colNum 列数
     * @return 表格对象
     */
    public static XWPFTable newTable(XWPFDocument doc, int rowNum, int colNum) {
        XWPFTable table = doc.createTable(rowNum, colNum);
        //设置指定宽度
        CTTbl ttbl = table.getCTTbl();
        CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
        CTTblLayoutType t = tblPr.isSetTblLayout()?tblPr.getTblLayout():tblPr.addNewTblLayout();
        t.setType(STTblLayoutType.FIXED);
        //表格宽度
        CTJc cTJc = tblPr.isSetJc()?tblPr.getJc():tblPr.addNewJc();
        cTJc.setVal(STJc.CENTER);   //表格居中

        table.setCellMargins(100,100,100,100);
        return table;
    }

    /**
     * 设置表格边框
     * @param table 表格
     * @param type 边框类型,默认传单线边框 XWPFTable.XWPFBorderType.SINGLE
     * @param size 边框大小, 一般传4
     * @param space 暂时不知道什么用
     * @param rgbColor 颜色, 可以传EFEFEF
     * @return 表格对象
     */
    public static XWPFTable setBorder(XWPFTable table, XWPFTable.XWPFBorderType type, int size, int space, String rgbColor) {
        //边框
//        CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
//        CTBorder hBorder = borders.addNewInsideH();
//        hBorder.setVal(STBorder.Enum.forString("single"));  // 线条类型
//        hBorder.setSz(BigInteger.valueOf(4)); // 线条大小
//        hBorder.setSpace(BigInteger.ZERO);  //space,暂时不清楚什么用
//        hBorder.setColor("EFEFEF"); // 设置颜色
//        hBorder.setNil();
        table.setTopBorder(type, size, space, rgbColor);
        table.setBottomBorder(type, size, space, rgbColor);
        table.setLeftBorder(type, size, space, rgbColor);
        table.setRightBorder(type, size, space, rgbColor);
        table.setInsideHBorder(type, size, space, rgbColor);
        table.setInsideVBorder(type, size, space, rgbColor);
        return table;
    }

    /**
     * 设置表格各列宽度
     * @param table 表格
     * @param weights 各列宽度所占权重
     */
    public static void setTableWidths(XWPFTable table, int tableWidth, int[] weights) {
        int[] colWidths = transWidth(tableWidth, weights);
        CTTbl ttbl = table.getCTTbl();
        CTTblGrid tblGrid = ttbl.addNewTblGrid();

        for (Integer i : colWidths) {
            CTTblGridCol gridCol = tblGrid.addNewGridCol();

            gridCol.setW(BigInteger.valueOf(i));
        }
    }

    /**
     * 设置跨页重复表头
     * <p>在各页顶端以标题形式重复出现
     * <p>实际就是在document.xml中增加以下行:
     * {@code
     *  <w:trPr>
     *      <w:tblHeader/>
     *  </w:trPr>
     * }
     * @param row 行对象
     */
    public static void setRepeatHeaderRow(XWPFTableRow row) {
        CTRow ctRow = row.getCtRow();
        CTTrPr trPr = ctRow.isSetTrPr() ? ctRow.getTrPr() : ctRow.addNewTrPr();
        trPr.addNewTblHeader();
//            cTOnOff.setVal(STOnOff.ON);
    }

    /**
     * 添加页眉(默认居右)
     * @param doc doc
     * @param consumer 自定义页眉
     * @return XWPFHeaderFooterPolicy
     */
    public static XWPFHeaderFooterPolicy header(XWPFDocument doc, Consumer<XWPFParagraph> consumer) {
        CTSectPr sectPr = DocxUtils.getPr(doc);
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(doc, sectPr);
        XWPFHeader header =  headerFooterPolicy.createHeader(STHdrFtr.DEFAULT);
        XWPFParagraph paragraph = header.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.RIGHT);
        paragraph.setVerticalAlignment(TextAlignment.CENTER);

        consumer.accept(paragraph);
        return headerFooterPolicy;
    }

    /**
     * 添加页眉(默认居右)
     * @param doc doc
     * @param text 页眉内容
     * @param fontFamily 字体名称,如:微 软 雅 黑
     * @param fontSize 字体大小
     * @param rgbStr 字体颜色(16进制)
     * @return 页眉引用
     */
    public static XWPFHeaderFooterPolicy header(XWPFDocument doc, String text, String fontFamily, int fontSize, String rgbStr) {
        return header(doc, (paragraph)-> addText(paragraph, text, fontFamily, fontSize, rgbStr));
    }

    /**
     * 添加页脚(默认居中)
     * @param doc doc
     * @param consumer 自定义页脚
     * @return XWPFHeaderFooterPolicy
     */
    public static XWPFHeaderFooterPolicy footer(XWPFDocument doc, Consumer<XWPFParagraph> consumer) {
        CTSectPr sectPr = getPr(doc);
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(doc, sectPr);
        XWPFFooter footer =  headerFooterPolicy.createFooter(STHdrFtr.DEFAULT);
        XWPFParagraph paragraph = footer.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setVerticalAlignment(TextAlignment.CENTER);

        consumer.accept(paragraph);
        return headerFooterPolicy;
    }

    /**
     * 添加页脚(默认居中)
     * {@code
     * footer(doc, "第", "页", 8, "000000") => 第X页
     * }
     * @param doc doc
     * @param before 页数前的字符串
     * @param after 页数后的字符串
     * @param fontFamily 字体名称,如:微 软 雅 黑
     * @param fontSize 字体大小
     * @param rgbStr 字体颜色(16进制)
     * @return 页脚引用
     */
    public static XWPFHeaderFooterPolicy footer(XWPFDocument doc, String before, String after, String fontFamily, int fontSize, String rgbStr) {
        return footer(doc, (paragraph)->{
            addText(paragraph, before, fontFamily, fontSize, rgbStr);
            addPageNum(paragraph, fontFamily, fontSize, rgbStr);
            addText(paragraph, after, fontFamily, fontSize, rgbStr);
        });
    }

    /**
     * 合并单元格
     * @param table table
     * @param row1 起点(左上角)行数
     * @param col1 起点列数
     * @param row2 终点(右下角)行数
     * @param col2 终点列数
     */
    public static void merge(XWPFTable table, int row1, int col1, int row2, int col2) {
        if(row1>row2 || col1>col2) {    //终止点小于起始点，异常，不合并
            return;
        }
        if(row1==row2 && col1==col2) {  //起止点一样，不需要合并
            return;
        }


        for (int i = row1; i <= row2; i++) {
            XWPFTableRow row = table.getRow(i);
            XWPFTableCell firstCol = row.getCell(col1);
            firstCol.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            if(row2>row1) {
                if(i == row1) {
                    firstCol.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
                } else {
                    firstCol.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
                }
            }
            for (int j = col1+1; j <= col2; j++) {
                row.getCell(j).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }




    }

    /**
     * 添加水印
     * @param doc doc
     * @param text 水印文本
     * @param colorHex 16进制色值
     * @param width 水印宽度
     * @param height 水印高度
     * @param rotationAngle 旋转角度
     */
    public static void addWaterMark(XWPFDocument doc, String text, String colorHex, int width, int height, int rotationAngle) {
        // the body content
        XWPFParagraph paragraph = doc.createParagraph();
        paragraph.createRun();
//        XWPFRun run=paragraph.createRun();
//        run.setText("The Body:");


        // create header-footer
        XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
        if (headerFooterPolicy == null) headerFooterPolicy = doc.createHeaderFooterPolicy();

        // create default Watermark - fill color black and not rotated
        headerFooterPolicy.createWatermark(text);

        //--
        // get the default header
        // Note: createWatermark also sets FIRST and EVEN headers
        // but this code does not updating those other headers
        XWPFHeader header = headerFooterPolicy.getHeader(XWPFHeaderFooterPolicy.DEFAULT);
        paragraph = header.getParagraphArray(0);

        // get com.microsoft.schemas.vml.CTShape where fill color and rotation is set
        org.apache.xmlbeans.XmlObject[] xmlobjects = paragraph.getCTP().getRArray(0).getPictArray(0).selectChildren(
                new javax.xml.namespace.QName("urn:schemas-microsoft-com:vml", "shape"));

        if (xmlobjects.length > 0) {
            com.microsoft.schemas.vml.CTShape ctshape = (com.microsoft.schemas.vml.CTShape)xmlobjects[0];
            // set fill color
            ctshape.setFillcolor(colorHex);
            // set rotation
            ctshape.setStyle(getWaterMarkStyle(ctshape.getStyle(),width, height) + ";rotation:"+rotationAngle);
            //System.out.println(ctshape);
        }
    }


    /**
     * 修改水印样式（水印的字体大小）
     * @param styleStr  水印的原样式
     * @param height    水印文字高度
     * @return
     */
    private static String getWaterMarkStyle(String styleStr, double width, double height){
        Pattern pattern=Pattern.compile("^(.*width:)(\\d+(?:\\.\\d+)?)pt(.*)height:(\\d+(?:\\.\\d+)?)pt(.*$)");
        Matcher matcher = pattern.matcher(styleStr);
        if(matcher.find()) {
//            System.out.println(matcher.group(1)+width+"pt"+matcher.group(3)+"height:"+height+"pt"+matcher.group(5));
            return matcher.group(1)+width+"pt"+matcher.group(3)+"height:"+height+"pt"+matcher.group(5);
        }
        return styleStr;
    }

    private static int getFormat(File file, int defaultValue) {
        String fileExtension = getFileExtension(file);
        switch(fileExtension) {
            case "emf":
                return XWPFDocument.PICTURE_TYPE_EMF;
            case "wmf":
                return XWPFDocument.PICTURE_TYPE_WMF;
            case "pict":
                return XWPFDocument.PICTURE_TYPE_PICT;
            case "jpeg":
            case "jpg":
                return XWPFDocument.PICTURE_TYPE_JPEG;
            case "png":
                return XWPFDocument.PICTURE_TYPE_PNG;
            case "dib":
                return XWPFDocument.PICTURE_TYPE_DIB;
            case "gif":
                return XWPFDocument.PICTURE_TYPE_GIF;
            case "tiff":
                return XWPFDocument.PICTURE_TYPE_TIFF;
            case "eps":
                return XWPFDocument.PICTURE_TYPE_EPS;
            case "bmp":
                return XWPFDocument.PICTURE_TYPE_BMP;
            case "wpg":
                return XWPFDocument.PICTURE_TYPE_WPG;
            default:
                return defaultValue;

        }
    }

    private static String getFileExtension(File file) {
        Pattern p = Pattern.compile("(?<=\\.)[^\\.]+$");
        Matcher m = p.matcher(file.getName());
        if(m.find()) {
            return m.group().toLowerCase();
        } else {
            return "";
        }
    }

    private static int[] transWidth(int tableWidth, int[] weights) {
        int sum = Arrays.stream(weights).sum();
        int[] widths = new int[weights.length];
        int total = 0;
        for (int i = 0; i < widths.length-1; i++) {
            int weight = weights[i];
            widths[i] = Math.round(weight*1f/sum*tableWidth);
            total+=widths[i];
        }
        widths[widths.length-1] = tableWidth-total;

        return widths;
    }

}
