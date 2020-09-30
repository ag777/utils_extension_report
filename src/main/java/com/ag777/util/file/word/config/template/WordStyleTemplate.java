package com.ag777.util.file.word.config.template;


import com.ag777.util.file.word.config.WordStyleInterf;

/**
 * docx表格样式模板
 *
 * @author ag777
 * @version create on 2020年09月30日,last modify at 2020年09月30日
 */
public class WordStyleTemplate implements WordStyleInterf {
    public static WordStyleTemplate mInstance;

    public static WordStyleTemplate getInstance() {
        if(mInstance == null) {
            synchronized (WordStyleTemplate.class) {
                if(mInstance == null) {
                    mInstance = new WordStyleTemplate();
                }
            }
        }
        return mInstance;
    }

    private static final int WIDTH = 11907;//(int)WordUtils.convert2Long(209f);
    private static final int HEIGHT = 16840;//(int)WordUtils.convert2Long(296f);

    private static final int MARGIN_LEFT_RIGHT = 1100;//WordUtils.convert2Long(19.4f);
    private static final int MARGIN_TOP_BOTTOM = 1220;//WordUtils.convert2Long(21.5f);

    private static final int WIDTH_TABLE = (int)(WIDTH-2*MARGIN_LEFT_RIGHT)-200;
    private static final String COLOR_BG_TABLE_TITLE = "f8f8f8";

    private static final int FONT_SIZE_TABLE_TITLE =9;
    private static final int FONT_SIZE_TABLE_CONTENT =9;

    private WordStyleTemplate() {}

    @Override
    public String fontFamily() {
        return "微 软 雅 黑";
    }

    @Override
    public int pageWidth() {
        return WIDTH;
    }

    @Override
    public int pageHeight() {
        return HEIGHT;
    }

    @Override
    public int pageMarginLeft() {
        return MARGIN_LEFT_RIGHT;
    }

    @Override
    public int pageMarginRight() {
        return MARGIN_LEFT_RIGHT;
    }

    @Override
    public int pageMarginTop() {
        return MARGIN_TOP_BOTTOM;
    }

    @Override
    public int pageMarginBottom() {
        return MARGIN_TOP_BOTTOM;
    }

    @Override
    public int tableWidth() {
        return WIDTH_TABLE;
    }

    @Override
    public String colorTableTitleBg() {
        return COLOR_BG_TABLE_TITLE;
    }

    @Override
    public int tableTitleFontSize() {
        return FONT_SIZE_TABLE_TITLE;
    }

    @Override
    public int tableContentFontSize() {
        return FONT_SIZE_TABLE_CONTENT;
    }

    @Override
    public int normalFontSize() {
        return FONT_SIZE_TABLE_CONTENT;
    }
}
