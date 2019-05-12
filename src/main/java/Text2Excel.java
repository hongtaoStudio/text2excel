import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author hongtao_studio@163.com
 * @date 2019/04/13
 * @description lua物编转为 Excel
 */
public class Text2Excel {

    private static LinkedHashMap<String, String> tableListName;

    private static ArrayList<Map<String, String>> itemList;

    //txt文本路径
    static String txtFilePath = "input.ini";
    //xls路径
    static String xlsFilePath = "output.xls";

    public static final Pattern PATTERN_ID = Pattern.compile("\\[(.*?)\\]");

    public static final Pattern PATTERN_NUM = Pattern.compile("^([-+])?\\d+(\\.\\d+)?$");

    public static void main(String args[]) {
        getTextInfos();
        transToExcel();
    }

    private static void getTextInfos() {
        tableListName = new LinkedHashMap<String, String>();
        itemList = new ArrayList<Map<String, String>>();
        String lastLine = "";

        BufferedReader bufferedReader = null;
        try {
            bufferedReader = new BufferedReader(new InputStreamReader(new FileInputStream(txtFilePath), "UTF-8"));
            String element;
            while ((element = bufferedReader.readLine()) != null) {
                if (element.trim().equals("")) {
                    continue;
                }
                Matcher m = PATTERN_ID.matcher(element);
                if (m.find() && element.length() == 6) {
                    Map<String, String> item = new HashMap<String, String>();
                    item.put("ID", m.group(1));
                    itemList.add(item);
                    tableListName.put("物编ID", "ID");
                } else if (element.startsWith("--")) {
                    lastLine = element.replace("-- ", "");
                } else {
                    Map<String, String> item = itemList.get(itemList.size() - 1);
                    String[] split = element.split(" = ");
                    String value = split[1];
                    if (value.startsWith("\"") && value.endsWith("\"")) {
                        value = value.substring(1, value.length() - 1);
                    }

                    item.put(split[0], value);
                    if (split[0].equals("_parent")) {
                        lastLine = "父对象";
                    }
                    tableListName.put(lastLine, split[0]);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (bufferedReader != null) {
                try {
                    bufferedReader.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private static void transToExcel() {
        WritableWorkbook book;
        try {
            book = Workbook.createWorkbook(new File(xlsFilePath));
            WritableSheet sheet = book.createSheet("英雄信息", 0);
            int column = 0;
            List<String> itemIndex = new ArrayList<String>();
            for (Map.Entry<String, String> set : tableListName.entrySet()) {
                String value = set.getValue();
                if (value.equals("_parent")) {
                    value = "parent";
                }
                sheet.addCell(new Label(column, 0, set.getKey()));
                sheet.addCell(new Label(column, 1, value));
                column++;
                itemIndex.add(set.getValue());
            }

            int row = 1;
            for (Map<String, String> itemMap : itemList) {
                row++;
                for (Map.Entry<String, String> item : itemMap.entrySet()) {
                    String key = item.getKey();
                    int index = itemIndex.indexOf(key);
                    if (key.endsWith("ID")) {
                        index = 0;
                    }
                    String value = item.getValue();
                    Matcher isNum = PATTERN_NUM.matcher(value);
                    if (value.equals("")) {
                        sheet.addCell(new Label(index, row, "#"));
                    } else if (!isNum.matches()) {
                        sheet.addCell(new Label(index, row, value));
                    } else {
                        sheet.addCell(new jxl.write.Number(index, row, Double.valueOf(value)));
                    }
                }
            }
            book.write();
            book.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
