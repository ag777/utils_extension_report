package com.ag777.util.file.chart;

import ChartDirector.*;
import com.ag777.util.lang.collection.ListUtils;
import com.ag777.util.lang.collection.MapUtils;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * 图表图片生成(二封ChartDirector)
 * <p>
 * 依赖jar包:
 * <ul>
 * <li>ChartDirector.jar</li>
 * </ul>
 *
 * @author ag777
 * @version create on 2020年09月30日,last modify at 2020年09月30日
 */
public class ChartBuildUtils {

    private static final String FONT_FAMILY1 = "simsun.ttc";
    private static final String FONT_FAMILY2 = "微软雅黑";
    private static final String FONT_FAMILY3 = "Arial";

//    static {
//        Chart.setLicenseCode("SXZVFNRN9MZ9L8LGA0E2B1BB");
//    }

    public static void main(String[] args) {
        barChartHorizon(
                "F:\\临时\\a.png",
                ListUtils.of(
                        MapUtils.of("key","数据1","size", 45),
                        MapUtils.of("key","数据2","size", 100)
                ), "key", "size", "尺寸", 0xc51313);
    }

    private ChartBuildUtils() {}

    public static File pieChart(String imgPath, double[] data, String[] labels, int[] colors){

        if(data.length == 0){//如果没有数据，则绘制一个空白小图片
            return emptyChart(imgPath);
        }

        int count = 0;
        for(double var: data){
            if(var > 0){
                count++;
            }
        }
        if(count == 0){//如果所有数据都是0，则绘制一个空白小图片
            PieChart c = new PieChart(1, 1);
            c.makeChart(imgPath);
            return new File(imgPath);
        }

//    	PieChart c = new PieChart(410, 250);
//    	c.setDefaultFonts(fontFamily);//设置字体
//        c.setPieSize(200, 100, 60);
        PieChart c = new PieChart(430, 200);
        c.setDefaultFonts(FONT_FAMILY1);//设置字体
        c.setPieSize(215, 100, 60);

        //c.set3D();

        c.setLabelLayout(Chart.SideLayout);

        c.setLineColor(Chart.SameAsMainColor, 0x000000);
        //c.setStartAngle(135);
        c.setLabelFormat("{label} {percent|1}%");
        c.setData(data, labels);

        for (int i = 0 ; i < colors.length ; i++) {
            TextBox t = c.sector(i).setLabelStyle("Microsoft YaHei", 9, 0x000000);//设置标注字体颜色
            //t.setBackground(Chart.SameAsMainColor, Chart.Transparent, Chart.glassEffect());
            //t.setBackground(Chart.SameAsMainColor, Chart.Transparent);
            t.setRoundedCorners(5);
            t.setFontColor(colors[i]);
            t.setFontStyle("underline");
            c.sector(i).setColor(colors[i]);
        }
        c.makeChart(imgPath);
        return new File(imgPath);
    }

    public static File barChart(String imgPath, double[][] data, int[] colors, String[] labels, String[] desc){
        if(data.length == 0){//如果没有数据，则绘制一个空白小图片
            return emptyChart(imgPath);
        }

        //XYChart c = new XYChart(507, 200);
        XYChart c = new XYChart(750, 270);

        c.setDefaultFonts(FONT_FAMILY1);//设置字体

        c.setPlotArea(40, 40, 690, 205,-1, 0xffffff, 0xffffff);

        c.addLegend(50, -5, false, FONT_FAMILY2, 8).setBackground(Chart.Transparent);
        String[] label = new String[]{"","","","",""};
        if(labels.length > 5){
            label = new String[labels.length];
        }
        System.arraycopy(labels, 0, label, 0, labels.length);

        c.xAxis().setLabels(label);

        c.xAxis().setTickOffset(0.5);

        c.xAxis().setLabelStyle(FONT_FAMILY2, 8);
        c.yAxis().setLabelStyle(FONT_FAMILY2, 8);

        c.xAxis().setWidth(2);
        c.yAxis().setWidth(2);

        BarLayer layer = c.addBarLayer2(Chart.Side, 4);
        for (int i = 0; i < desc.length; i++) {
            layer.addDataSet(data[i],colors[i],desc[i]);
        }

        layer.setAggregateLabelStyle();

        layer.setBorderColor(Chart.Transparent, Chart.barLighting(0.75, 1.75));

        layer.setBarGap(0.2, Chart.TouchBar);

        c.packPlotArea(30, 25, c.getWidth() - 50, c.getHeight() - 25);
        c.makeChart(imgPath);
        return new File(imgPath);
    }

    public static File barChart(String imgPath, double[][] data, int[] colors, String[] labels, String[] desc, int xNum){
        if(data.length == 0){//如果没有数据，则绘制一个空白小图片
            return emptyChart(imgPath);
        }

        XYChart c = new XYChart(750, 270);

        c.setDefaultFonts(FONT_FAMILY1);//设置字体

        c.setPlotArea(40, 40, 690, 205,-1, 0xffffff, 0xffffff);

        c.addLegend(50, -5, false, FONT_FAMILY2, 8).setBackground(Chart.Transparent);
        String[] label = new String[xNum];
        System.arraycopy(labels, 0, label, 0, labels.length);
        if(labels.length < xNum){//若 labels 不足 xNum 个，则将 label 补足到 xNum 个
            for(int j=labels.length; j < xNum; j++){
                label[j] = "";
            }
        }

        c.xAxis().setLabels(label);

        c.xAxis().setTickOffset(0.5);

        c.xAxis().setLabelStyle(FONT_FAMILY2, 8);
        c.yAxis().setLabelStyle(FONT_FAMILY2, 8);

        c.xAxis().setWidth(2);
        c.yAxis().setWidth(2);

        BarLayer layer = c.addBarLayer2(Chart.Side, 4);
        for (int i = 0; i < desc.length; i++) {
            layer.addDataSet(data[i],colors[i],desc[i]);
        }

        layer.setAggregateLabelStyle();

        layer.setBorderColor(Chart.Transparent, Chart.barLighting(0.75, 1.75));

        layer.setBarGap(0.2, Chart.TouchBar);

        c.packPlotArea(30, 25, c.getWidth() - 50, c.getHeight() - 25);
        c.makeChart(imgPath);
        return new File(imgPath);
    }

    public static File barChartHorizon(String imgPath, List<Map<String, Object>> dataList, boolean removeLast,
                                       String[] legends, String[] attrNames, int[] colors,
                                       String xAttrName){
        if(dataList.size() == 0){//如果没有数据，则绘制一个空白小图片
            return emptyChart(imgPath);
        }

        //int[] colors = NnvasSystem.vulColors;
        //int[] colors = {0x71ddcd, 0x68b8f0, 0xfcb45b, 0xfa7978, 0xc31512};//信息->紧急

        //XYChart c = new XYChart(507, 200);
        //XYChart c = new XYChart(750, 270);
        int size = dataList.size();
        if(removeLast && size > 1){//将最后一行数据排除
            size = size - 1;
        }

//        	if(num > 50){
//        		num = 50;
//        	}
        XYChart c = newXYChart(size);

        //x轴label
        String[] xLabels = new String[size];
        for(int n = 0; n < size ; n++){
            xLabels[n] = MapUtils.getStr(dataList.get(size-n-1), xAttrName, "");
        }
        c.xAxis().setLabels(xLabels);
        c.yAxis().setTickDensity(40);

        c.xAxis().setLabelStyle(FONT_FAMILY2, 8);
        c.yAxis().setLabelStyle(FONT_FAMILY2, 8);

        //c.syncYAxis();

        //y轴数据
        //BarLayer layer = c.addBarLayer2(Chart.Side, 4);
        BarLayer layer = c.addBarLayer2(Chart.Stack);
        for(int i = 0; i < legends.length; i++){
            double[] data = new double[size];
            for(int n = 0; n < size ; n++){
                data[n] = MapUtils.getDouble(dataList.get(size-n-1), attrNames[i], 0);
            }
            layer.addDataSet(data, colors[i], legends[i]);
        }
        //layer.setLegendOrder(Chart.ReverseLegend);
        layer.setDataLabelStyle(FONT_FAMILY3, 10).setAlignment(Chart.Center);

        layer.setAggregateLabelStyle(FONT_FAMILY3, 12);

        layer.setBorderColor(Chart.Transparent);
        layer.setBarGap(0.35); //设置柱状间距

        c.packPlotArea(30, 25, c.getWidth() - 50, c.getHeight() - 25);
        c.makeChart(imgPath);

        return new File(imgPath);
    }

    public static File barChartHorizon(String imgPath, List<Map<String, Object>> dataList, String keyKey, String valueKey, String legendStr, int colorInt){

        if(dataList.size() == 0){//如果没有数据，则绘制一个空白小图片
            return emptyChart(imgPath);
        }

        //int[] colors = NnvasSystem.vulColors;
        //int[] colors = {0x71ddcd, 0x68b8f0, 0xfcb45b, 0xfa7978, 0xc31512};//信息->紧急

        //XYChart c = new XYChart(507, 200);
        //XYChart c = new XYChart(750, 270);
        int size = dataList.size();

//        	if(num > 50){
//        		num = 50;
//        	}
        XYChart c = newXYChart(size);

        //x轴设置
        String[] xLabels = new String[size];
        c.yAxis().setTickDensity(40);

        c.xAxis().setLabelStyle(FONT_FAMILY2, 8);
        c.yAxis().setLabelStyle(FONT_FAMILY2, 8);

        //c.syncYAxis();

        //y轴设置
        BarLayer layer = c.addBarLayer2(Chart.Stack);
        double[] data = new double[size];

        //layer.setLegendOrder(Chart.ReverseLegend);
        layer.setDataLabelStyle(FONT_FAMILY3, 10).setAlignment(Chart.Center);

        layer.setAggregateLabelStyle(FONT_FAMILY3, 12);

        layer.setBorderColor(Chart.Transparent);
        layer.setBarGap(0.35); //设置柱状间距

        /*填充数据*/
        for(int i = 0; i < size ; i++){
            Map<String, Object> dataMap = dataList.get(i);
            String key = MapUtils.getStr(dataMap, keyKey, "");
            double value = MapUtils.getDouble(dataMap, valueKey, 0);
            int j = size-i-1;	//倒着差，可以让图表数据从上到下顺序
            xLabels[j] = key;
            data[j] = value;
        }

        c.xAxis().setLabels(xLabels);
//          c.yAxis().setLabels(IntStream.range(0, 120).filter(i->i%10==0).asDoubleStream().toArray());
        layer.addDataSet(data, colorInt, legendStr);

        c.packPlotArea(30, 25, c.getWidth() - 50, c.getHeight() - 25);
        c.makeChart(imgPath);
        return new File(imgPath);
    }


    public static File emptyChart(String imgPath) {
        PieChart c = new PieChart(1, 1);
        c.makeChart(imgPath);
        return new File(imgPath);
    }

    private static XYChart newXYChart(int size) {
        XYChart c = new XYChart(750, 100 + 32*size);
        c.xAxis().setColors(Chart.Transparent);
        c.yAxis().setColors(Chart.Transparent);

        c.setDefaultFonts(FONT_FAMILY1);//设置字体

        //c.setPlotArea(20, 20, 467, 150,-1, 0xffffff, 0xffffff);
        //c.setPlotArea(40, 40, 690, 205,-1, 0xffffff, 0xffffff);
        c.setPlotArea(70, 80, 690, 80 + 28*size, Chart.Transparent, -1, Chart.Transparent, 0xcccccc);

        c.swapXY();

        c.addLegend(50, 0, false, FONT_FAMILY2, 10).setBackground(Chart.Transparent);

        return c;
    }

}
