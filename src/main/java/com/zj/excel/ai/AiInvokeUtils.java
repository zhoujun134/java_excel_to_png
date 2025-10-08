package com.zj.excel.ai;

import cn.hutool.ai.AIServiceFactory;
import cn.hutool.ai.ModelName;
import cn.hutool.ai.core.AIConfig;
import cn.hutool.ai.core.AIConfigBuilder;
import cn.hutool.ai.model.openai.OpenaiService;
import cn.hutool.core.date.DateUtil;
import cn.hutool.json.JSONObject;
import cn.hutool.json.JSONUtil;
import com.zj.excel.ai.domain.ChatCompletion;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import java.util.Objects;

/**
 * @ClassName AiService
 * @Author zj
 * @Description
 * @Date 2025/10/1 16:20
 * @Version v1.0
 **/
@Slf4j
public class AiInvokeUtils {
    private static AIConfig aiConfig = new AIConfigBuilder(ModelName.OPENAI.getValue())
            .setApiKey(".....")
            .setApiUrl("http://localhost:11434/v1")
            .setModel("gpt-oss:20b-cloud")
//            .setModel("qwen3:0.6b")
            .build();

    public void setAiConfig(AIConfig aiConfig) {
        if (Objects.isNull(aiConfig)) {
            return;
        }
        AiInvokeUtils.aiConfig = aiConfig;
    }

    private final static OpenaiService openaiService = AIServiceFactory.getAIService(aiConfig, OpenaiService.class);

    public static String invoke(String prompt) {
        long startTime = System.currentTimeMillis();
        String startTimeStr = DateUtil.now();
        String chatResult = "";
        try {
            chatResult = openaiService.chat(prompt);
            JSONObject jsonObject = JSONUtil.parseObj(chatResult);
            ChatCompletion chatCompletion = jsonObject.toBean(ChatCompletion.class);
            String content = chatCompletion.getChoices().get(0).getMessage().getContent();
            if (StringUtils.isNotBlank(content)) {
                // 找出 <think> </think> 标签之间的内容（包含标签）替换为空格
                content = content.replaceAll("(?s)<think>.*?</think>", "").trim();
            }
            return content;
        } catch (Exception exception) {
            log.error("调用 ai 错误，错误信息: {}", exception.getMessage());
            log.warn("出错的返回结果为: {}", chatResult);
        } finally {
            long endTime = System.currentTimeMillis();
            String endTimeStr = DateUtil.now();
            long cost = endTime - startTime;
            log.info("invoke 调用 ai 结束，开始时间: {}, 结束时间，{} 整体耗时: cost:{}", startTimeStr, endTimeStr, cost);
        }
        return chatResult;
    }
    public static <T> T invokeWithJson(String prompt, Class<T> clazz) {
        return baseInvokeWithJson(prompt, clazz, 0);
    }

    public static <T> T baseInvokeWithJson(String prompt, Class<T> clazz, int retryCount) {
        long startTime = System.currentTimeMillis();
        String startTimeStr = DateUtil.now();
        String chatResult = "";
        try {
            if (retryCount >= 3) {
                log.warn("重试了 3 次了仍然没有获取到相关数据，直接返回 null 了");
                return null;
            }
            chatResult = invoke(prompt);
            boolean isJson = JSONUtil.isTypeJSON(chatResult);
            if (isJson) {
                return JSONUtil.toBean(chatResult, clazz);
            }
            String promptJson = String.format("你返回的结果不是一个 json 字符串，请你去掉非json字符串的内容，只返回 json的结果即可。原始结果为: %s", chatResult);
            return baseInvokeWithJson(promptJson, clazz, retryCount + 1);
        } catch (Exception exception) {
            log.error("invokeWithJson 调用 ai 错误，错误信息: {}", exception.getMessage());
            log.warn("invokeWithJson 出错的返回结果为: {}", chatResult);
            return null;
        } finally {
            long endTime = System.currentTimeMillis();
            String endTimeStr = DateUtil.now();
            long cost = endTime - startTime;
            log.info("invokeWithJson 调用 ai 结束，开始时间: {}, 结束时间，{} 整体耗时: cost:{}", startTimeStr, endTimeStr, cost);
        }
    }
}

