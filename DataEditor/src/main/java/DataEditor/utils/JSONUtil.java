package DataEditor.utils;

import com.alibaba.fastjson.JSON;

/**
 * Created by muskong on 2016/8/25.
 */
public class JSONUtil {
    public static final String toJSONString(Object object){
        return JSON.toJSONString(object);
    }

//    public static final <T> T parseObject(String text){
//        return JSON.parseObject(text);
//    }

}
