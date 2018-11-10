package com.gzcdc.officeserver.util;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by User: admin.
 * Date: 2017/10/27 Time: 17:06.
 * Description:
 */
public class DateUtil {
    /**
     * 获取时间
     *
     * @return
     */
    public static String getNowDate() {
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");// 设置日期格式
        return df.format(new Date());
    }
}
