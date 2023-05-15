import com.ag777.util.file.word.DocxBuilder;
import com.ag777.util.file.word.DocxUtils;
import com.ag777.util.file.word.config.TableItemConfig;
import com.ag777.util.file.word.config.template.WordStyleTemplate;
import com.ag777.util.lang.collection.ListUtils;
import com.ag777.util.lang.collection.MapUtils;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * word导出示例
 * @author ag777 <837915770@vip.qq.com>
 * @version 2023/5/12 17:56
 */
public class WordExportDemo {
    public static void main(String[] args) throws IOException {
        // 数据列表
        List<Map<String, Object>> dataList = IntStream.range(1, 10).boxed().map(i->MapUtils.of(
                "a", "a_"+i,
                "b", "b_"+i,
                "c", "c_"+i
        )).collect(Collectors.toList());
        // 行配置列表
        List<TableItemConfig<Map<String, Object>>> configList = ListUtils.of(
                new TableItemConfig<Map<String, Object>>()
                        .setTitle("第一列")
                        .setKey("a")
                        .setContentRender((cell, item, key) -> {
                            cell.setText(MapUtils.getStr(item, key));
                        }),
                new TableItemConfig<Map<String, Object>>()
                        .setTitle("第二列")
                        .setKey("b")
                        .setAlign(STJc.RIGHT)
                        .setContentRender((cell, item, key) -> cell.setText(MapUtils.getStr(item, key))),
                new TableItemConfig<Map<String, Object>>()
                        .setTitle("第三列")
                        .setKey("c")
                        .setAlign(STJc.RIGHT)
                        .setContentRender((cell, item, key) -> cell.setText(MapUtils.getStr(item, key)))
        );
        // 构建器
        DocxBuilder builder = new DocxBuilder(WordStyleTemplate.getInstance());
        // 添加表格
        XWPFTable table = builder
                .addTable(dataList, configList, true);
        // 合单元格, 从(1,0)到(2,1),矩形,除了合并原点，其它数据会被隐藏
        DocxUtils.merge(table, 1, 0, 2, 1);
        // 加边框
        DocxUtils.setBorder(table, XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        // 输出到文件
        builder.save(new File("d:/c.docx"));
    }


}
