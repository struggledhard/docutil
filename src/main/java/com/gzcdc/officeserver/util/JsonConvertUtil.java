package com.gzcdc.officeserver.util;

import com.google.gson.JsonElement;
import com.google.gson.JsonParser;
import fr.opensagres.xdocreport.document.json.JSONArray;
import fr.opensagres.xdocreport.document.json.JSONObject;

import java.util.*;

/**
 * Created by User: admin.
 * Date: 2018/5/17 Time: 16:12.
 * Description:json转map
 */
public class JsonConvertUtil {
    public static List<Map<String, Object>> jsonConvert(String jsonStr){
        List<Map<String,Object>> list = new ArrayList<>();
        if(jsonStr.startsWith("[")){
            com.google.gson.JsonArray jarray =  new JsonParser().parse(jsonStr).getAsJsonArray();
            Iterator<com.google.gson.JsonElement> it = jarray.iterator();
            while (it.hasNext()) {
                JsonElement json2 = it.next();
                list.add(jsonToMap(json2.toString()));
            }
        }else  if(jsonStr.startsWith("{")){
            list.add(jsonToMap(jsonStr));
        }
        return list;
    }

    /**
     * 将json字符串转为Map结构
     * 如果json复杂，结果可能是map嵌套map
     * @param jsonStr 入参，json格式字符串
     * @return 返回一个map
     */
    public static Map<String, Object> jsonToMap(String jsonStr) {
        Map<String, Object> map = new HashMap<>();
        if (jsonStr != null && !"".equals(jsonStr)) {
            // 解析最外层
            JSONObject jsonObject = new JSONObject(jsonStr);
            for (Object k : jsonObject.keySet()) {
                Object v = jsonObject.get(k);
                if (v instanceof JSONArray) {
                    List<Map<String,Object>> list = new ArrayList<>();
                    Iterator<JSONObject> it = ((JSONArray) v).iterator();
                    while (it.hasNext()) {
                        JSONObject json2 = it.next();
                        list.add(jsonToMap(json2.toString()));
                    }
                    map.put(k.toString(), list);
                } else {
                    map.put(k.toString(), v);
                }
            }
        }
        return map;
    }
}
