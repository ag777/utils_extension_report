package com.ag777.util.file.word.config;

import com.ag777.util.lang.function.TriConsumer;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

import java.util.function.BiConsumer;
import java.util.function.BiFunction;

/**
 * 表格项配置
 *
 * @author ag777
 * @version create on 2020年09月30日,last modify at 2023年05月15日
 */
public class TableItemConfig<T> {
    private String title;
    private String key;
    private int weight;
    /** 可选,水平对齐模式,默认为STJc.CENTER */
    private STJc.Enum align;
    /** 可选,16进制字体色值 */
    private String colorHex;
    /** 渲染标题, (cell, 标题) */
    private BiConsumer<XWPFTableCell, String> titleRender;
    /** 直接获取文字自动渲染,(数据项, key, 展示内容) */
    private BiFunction<T, String, String> getContent;
    /** 完全自定义展示,(cell, 数据项, key) *getContent不设置时执行 */
    private TriConsumer<XWPFTableCell, T, String> contentRender;

    public TableItemConfig() {
        this.weight = 1;
        this.align = STJc.CENTER;
    }

    public TableItemConfig(String title, String key) {
        this();
        this.title = title;
        setKey(key);
    }

    public String getTitle() {
        return title;
    }

    public TableItemConfig<T> setTitle(String title) {
        this.title = title;
        return this;
    }

    public String getKey() {
        return key;
    }

    public TableItemConfig<T> setKey(String key) {
        this.key = key;
        return this;
    }

    public int getWeight() {
        return weight;
    }

    public TableItemConfig<T> setWeight(int weight) {
        this.weight = weight;
        return this;
    }

    public STJc.Enum getAlign() {
        return align;
    }

    public TableItemConfig<T> setAlign(STJc.Enum align) {
        this.align = align;
        return this;
    }

    public String getColorHex() {
        return colorHex;
    }

    public TableItemConfig setColorHex(String colorHex) {
        this.colorHex = colorHex;
        return this;
    }

    public BiConsumer<XWPFTableCell, String> getTitleRender() {
        return titleRender;
    }

    public TableItemConfig setTitleRender(BiConsumer<XWPFTableCell, String> titleRender) {
        this.titleRender = titleRender;
        return this;
    }

    public BiFunction<T, String, String> getGetContent() {
        return getContent;
    }

    public TableItemConfig setGetContent(BiFunction<T, String, String> getContent) {
        this.getContent = getContent;
        return this;
    }

    public TriConsumer<XWPFTableCell, T, String> getContentRender() {
        return contentRender;
    }

    public TableItemConfig setContentRender(TriConsumer<XWPFTableCell, T, String> contentRender) {
        this.contentRender = contentRender;
        return this;
    }
}
