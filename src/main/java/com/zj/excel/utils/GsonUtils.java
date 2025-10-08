package com.zj.excel.utils;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.reflect.TypeToken;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Type;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

@Slf4j
public class GsonUtils {

    private final static Gson gson = new GsonBuilder()
            .disableHtmlEscaping()
            .registerTypeAdapter(LocalDateTime.class, new LocalDateTimeAdapter())
            .create();

    private final static ObjectMapper objectMapper = new ObjectMapper();

    public static DateTimeFormatter format = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    public static String toJSONString(Object data) {
        if (Objects.isNull(data)) {
            return "";
        }
        try {
            return objectMapper.writeValueAsString(data);
//            return gson.toJson(data);
        } catch (Exception e) {
            log.error("JsonUtils######toString： error data={}", data, e);
            return "";
        }
    }

    public static <T> T parseObject(String text, Class<T> tClass) {
        if (StringUtils.isBlank(text)) {
            return null;
        }
        try {
            return gson.fromJson(text, tClass);
        } catch (Exception e) {
            log.error("JsonUtils######parseObject： error data={}", text, e);
            return null;
        }
    }

    public static <T> List<T> parseArray(String text, Class<T> clazz) {
        List<T> result = Collections.emptyList();
        try {
            if (StringUtils.isBlank(text)) {
                text = "[]";
            }
            result = gson.fromJson(text, TypeToken.getParameterized(ArrayList.class, new Type[]{clazz}).getType());
        } catch (Exception e) {
            String message =
                    String.format("GsonUtil######parseObject 对象转换 json 异常! text=%s, clazz=%s", text, clazz);
            log.error(message, e);
            throw new RuntimeException(message);
        }
        return result;
    }

    public static <K, V> Map<K, V> parseMap(String text) {
        Map<K, V> result = Collections.emptyMap();
        try {
            if (StringUtils.isBlank(text)) {
                text = "{}";
            }
            result = gson.fromJson(text, new TypeToken<Map<K, V>>() {
            }.getType());
        } catch (Exception e) {
            String message = String.format("GsonUtil######parseMap 对象转换 map 异常! text=%s", text);
            log.error(message, e);
            throw new RuntimeException(message);
        }
        return result;
    }
}
