package cn.myccit.ifactory.action.utils;


/**
 * TODO
 *
 * @Author jack
 * @Since 1.0 2020/3/18 10:06
 */

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.*;

/**
 * 读取word文档中表格数据，支持doc、docx
 *
 * @author Fise19
 */
public class ReadBomDocUtils {
    public static void main(String[] args) {

        //String filePath = "C:\\Users\\jack\\Desktop\\远程\\生产BOM\\法半夏20200201.doc";
        String filePath = "C:\\Users\\jack\\Desktop\\远程\\生产BOM\\紫菀20200201.doc";
        try {
            FileInputStream in = new FileInputStream(filePath);
            //载入文档
            JSONObject readWord = readWord(in);
            System.out.println(readWord.toString());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

    }


    /**
     * 获取文档特定字符的特定位置
     * @param tb
     * @return
     */
    private static Integer getDocIndex(Table tb,String positionStr) {
        Integer index = null;
        for (int i = 0; i < tb.numRows(); i++) {
            String str = tb.getRow(i).getCell(0).getParagraph(0).text().trim();
            if (positionStr.equals(str)) {
                index = i;
                break;
            }
        }
        return index;
    }


    /**
     * 读取文档中表格
     *
     * @param in
     */
    public static JSONObject readWord(InputStream in) {
        String key = null;//表头key、value
        String value = null;
        String content;//单元格内容
        Map<String, String> mapTop = new HashMap<>();//表头
        List<List<String>> operationData = new ArrayList<>(); // 工序标准详情
        List<List<String>> itemData = new ArrayList<>(); // 物料详情
        List<String> item = new ArrayList<>();
        String router = ""; // 工艺路线
        String remark = ""; // 备注
        JSONObject mainData = null; // 返回的解析数据
        try {
            // 处理doc格式 即office2003版本
            POIFSFileSystem pfs = new POIFSFileSystem(in);
            HWPFDocument hDoc = new HWPFDocument(pfs);
            Range range = hDoc.getRange();//得到文档的读取范围
            TableIterator it = new TableIterator(range);

            while (it.hasNext()) {
                Table tb = (Table) it.next();
                Integer blankIndex = getDocIndex(tb,"——"); // 获取—— 的位置
                Integer techIndex = getDocIndex(tb,"工艺规程"); // 获取—— 的位置
                Integer commandUserIndex = getDocIndex(tb,"指令编制人"); // 获取—— 的位置
                if (blankIndex == null) {
                    break;
                }else {
                    // 读取中间的表格
                    for (int i = 0; i < tb.numRows(); i++) {  // 11行
                        TableRow tr = tb.getRow(i);
                        //迭代列，默认从0开始
                        for (int j = 0; j < tr.numCells(); j++) {
                            TableCell td = tr.getCell(j);//取得单元格
                            //取得单元格的内容
                            int t = 0;
                            for (int k = 0; k < td.numParagraphs(); k++) {

                                Paragraph para = td.getParagraph(k);
                                content = para.text();
                                //去除后面的特殊符号
                                if (null != content && !"".equals(content)) {
                                    content = content.substring(0, content.length() - 1).trim();
                                }
                                //System.out.println(String.format("%s, %s,%s,%s",content,i,j,k));
                                if (para.getTableLevel() == 2) {
                                    if (para.isTableRowEnd()) {
                                        if (k-t == 6 || k-t == 1 ) {
                                            List<String> arr = new ArrayList<>();
                                            arr.add(getHeaderContent(k-5,td));
                                            arr.add(getSecondContent(k-4,td));
                                            arr.add(getSecondContent(k-3,td));
                                            arr.add(getSecondContent(k-2,td));
                                            arr.add(getSecondContent(k-1,td));
                                            operationData.add(arr);
                                        }else if (k-t == 8){
                                            // 添加两次
                                            List<String> arr1 = new ArrayList<>();
                                            arr1.add(td.getParagraph(k-7).text().trim());
                                            arr1.add(td.getParagraph(k-6).text().trim());
                                            arr1.add(td.getParagraph(k-5).text().trim());
                                            arr1.add(td.getParagraph(k-3).text().trim());
                                            arr1.add(td.getParagraph(k-1).text().trim());
                                            operationData.add(arr1);
                                            List<String> arr2 = new ArrayList<>();
                                            arr2.add(td.getParagraph(k-7).text().trim());
                                            arr2.add(td.getParagraph(k-6).text().trim());
                                            arr2.add(td.getParagraph(k-4).text().trim());
                                            arr2.add(td.getParagraph(k-2).text().trim());
                                            arr2.add(td.getParagraph(k-1).text().trim());
                                            operationData.add(arr2);
                                        }else if (k-t == 9){
                                            // 添加两次
                                            List<String> arr1 = new ArrayList<>();
                                            arr1.add(td.getParagraph(k-8).text().trim());
                                            arr1.add(td.getParagraph(k-7).text().trim());
                                            arr1.add(td.getParagraph(k-5).text().trim());
                                            arr1.add(td.getParagraph(k-3).text().trim());
                                            arr1.add(td.getParagraph(k-1).text().trim());
                                            operationData.add(arr1);
                                            List<String> arr2 = new ArrayList<>();
                                            arr2.add(td.getParagraph(k-8).text().trim());
                                            arr2.add(td.getParagraph(k-7).text().trim());
                                            arr2.add(td.getParagraph(k-4).text().trim());
                                            arr2.add(td.getParagraph(k-2).text().trim());
                                            arr2.add(td.getParagraph(k-1).text().trim());
                                            operationData.add(arr2);
                                        }else if (k-t == 7) {
                                            List<String> arr3 = new ArrayList<>();
                                            // 判断是否有分号结尾
                                            String t3 = td.getParagraph(k-3).text().trim();
                                            if (t3.contains("第一")) {
                                                arr3.add(getHeaderContent(k-6,td));
                                                arr3.add(getSecondContent(k-5,td));
                                                arr3.add(getSecondContent(k-4,td));
                                                arr3.add(getSecondContent(k-3,td)+getSecondContent(k-2,td));
                                                arr3.add(getSecondContent(k-1,td));
                                            }else {
                                                arr3.add(getSecondContent(k-6,td));
                                                arr3.add(getSecondContent(k-5,td));
                                                arr3.add(getSecondContent(k-3,td));
                                                arr3.add(getSecondContent(k-2,td));
                                                arr3.add(getSecondContent(k-1,td));
                                            }
                                            operationData.add(arr3);
                                        }
                                        t= k;
                                    }
                                }else if(para.getTableLevel() == 1 ){
                                    if (i == 0 && i< (blankIndex-1)) {
                                        // 取表头
                                        if (j == 0 || j % 2 == 0) {
                                            key = content;
                                        } else {
                                            value = content;
                                        }
                                        mapTop.put(key, value);
                                        if (mapTop.containsKey("产日期")) {
                                            mapTop.put("计划生产日期",tb.getRow(i).getCell(7).getParagraph(0).text().trim());
                                        }
                                    }else if (i >0  && i < blankIndex) {
                                        // 取工序
                                        item.add(content);
                                        if( j == 3) {
                                            itemData.add(item);
                                            item = new ArrayList<>();
                                        }
                                    }else if (i == techIndex && j == 1) {
                                        // 取工艺规程
                                        mapTop.put("工艺规程", content);
                                    }else if (i == techIndex+1 && j > 0) {
                                        // 取工艺路线
                                        router = router + content+"\n";
                                    }else if (i == commandUserIndex){
                                        if (j == 0 || j % 2 == 0) {
                                            key = content.contains("日期") ? "指令单日期" : content;
                                        } else {
                                            value = content;
                                        }
                                        mapTop.put(key, value);
                                    }else if (i == commandUserIndex+1){
                                        if (j == 0 || j % 2 == 0) {
                                            key = content.contains("日期") ? "批准日期" : content;
                                        } else {
                                            value = content;
                                        }
                                        mapTop.put(key, value);
                                        //System.out.println(String.format("%s, %s,%s,%s",content,i,j,k));
                                    }else if(i == commandUserIndex+2 && k > 0) {
                                        // 备注
                                        remark = remark + content + "\n";
                                    }
                                }
                            }
                        }
                    }
                }

            }
            mapTop.put("工艺路线",router);
            mapTop.put("备注",remark);
            // 设置主要信息
            mainData = setMainData(mapTop);
            // 设置物料信息
            JSONArray itemArray = setItemData(itemData);
            mainData.put("item",itemArray);
            // 设置工序
            //System.err.println(operationData);
            JSONArray operation = setOperationData(operationData);
            mainData.put("operation",operation);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return mainData;
    }



    /**
     * 一级
     * @param k
     * @param td
     * @return
     */
    private static String getHeaderContent(int k, TableCell td){
        String res;
        String t1 = td.getParagraph(k).text().trim();
        if ("".equals(t1)) {
            String t3 = td.getParagraph(k-8).text().trim();
            String t5 = td.getParagraph(k-7).text().trim();
            if (!"".equals(t5) || ("".equals(t5) && "".equals(t3))){
                res =  getHeaderContent(k-7,td);
            }else {
                res = getHeaderContent(k-6,td);
            }
        }else {
            res = td.getParagraph(k).text().trim();
        }

        return res;
    }

    /**
     * 二级
     * @param k
     * @param td
     * @return
     */
    private static String getSecondContent(int k, TableCell td){
        String res;
        String t1 = td.getParagraph(k).text().trim();
        if ("".equals(t1)) {
            res =  getSecondContent(k-6,td);
        }else {
            res = td.getParagraph(k).text().trim();
        }
        return res;
    }

    /**
     * 设置主要信息
     * @param mapTop
     * @return
     */
    private static JSONObject setMainData(Map<String, String> mapTop) {
        JSONObject res = new JSONObject();
        // 表头信息
        res.put("productName", mapTop.get("产品名称"));
        res.put("productBno", mapTop.get("产品批号"));
        res.put("productNo", mapTop.get("产品编码"));
        res.put("productTime", mapTop.get("计划生产日期"));
        res.put("techniqueRouter",mapTop.get("工艺路线"));
        res.put("techniqueStandard",mapTop.get("工艺规程"));
        res.put("commandUser",mapTop.get("指令编制人"));
        res.put("commandTime",mapTop.get("指令单日期"));
        res.put("approveUser",mapTop.get("批准人"));
        res.put("approveTime",mapTop.get("批准日期"));
        res.put("remark",mapTop.get("备注"));
        return res;
    }

    /**
     * 设置物料信息  ——
     * @param itemList
     * @return
     */
    private static JSONArray setItemData(List<List<String>> itemList){
        JSONArray itemArray = new JSONArray();
        for(int x=1; x< itemList.size(); x++) {
            JSONObject obj = new JSONObject();
            List<String> strList = itemList.get(x);
            obj.put("itemName", strList.get(0));
            obj.put("itemBno", strList.get(1));
            obj.put("itemNo", strList.get(2));
            obj.put("itemNum", strList.get(3));
            itemArray.add(obj);
        }
        return itemArray;
    }

    /**
     * 设置工序信息
     * @param operationList
     * @return
     */
    private static JSONArray setOperationData(List<List<String>> operationList){
        JSONArray operation = new JSONArray();
        for(int x=0; x< operationList.size(); x++) {
            JSONObject obj = new JSONObject();
            List<String> strList = operationList.get(x);
            if (!"工序".equals(strList.get(0))) {
                obj.put("operation", strList.get(0));
                obj.put("monitor", strList.get(1));
                obj.put("project", strList.get(2));
                obj.put("standard", strList.get(3));
                obj.put("freq", strList.get(4));
                operation.add(obj);
            }
        }
        return operation;
    }
}


