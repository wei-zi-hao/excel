package com.ek.project.common;

import org.springframework.stereotype.Component;

import java.util.HashMap;
import java.util.Map;

@Component
public class GlobalHashMapHolder
{
    // 创建一个全局的HashMap
    private Map<String, Object> globalHashMap = new HashMap<>();

    // 提供一个获取全局HashMap的方法
    public Map<String, Object> getGlobalHashMap() {
        return globalHashMap;
    }

    // 提供一个设置值的方法（可选，根据实际需求）
    public void putIntoGlobalHashMap(String key, Object value) {
        this.globalHashMap.put(key, value);
    }
}
