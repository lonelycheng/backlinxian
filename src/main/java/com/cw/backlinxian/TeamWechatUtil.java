package com.cw.backlinxian;

import com.alibaba.fastjson.JSONObject;
import org.apache.commons.lang.StringUtils;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.StringEntity;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.entity.mime.content.FileBody;
import org.apache.http.entity.mime.content.StringBody;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;

import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;

/**
 * 企业微信工具类
 */
public class TeamWechatUtil {

    // 企业微信测试企业账号的id和 防疫小应用的secret
    public static final String CORPID = "ww5d6d97ad43b9170c";
    public static final String CORPSECRET = "6ULyZwMZRUQI1zRzs7Oj1wnmhHCxSfpvlCXNF1lcf58";
    public static final String USERID = "ChengWei"; // 需要发送给谁信息
    public static final int AGENTID = 1000002; // 需要发送给谁信息

    static CloseableHttpClient httpclient = HttpClients.createDefault();

    public static String getAccessToken() {
        String accessToken = "";
        HttpGet httpGet = new HttpGet("https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid="+CORPID+"&corpsecret=" + CORPSECRET);
        try {
            HttpResponse response = httpclient.execute(httpGet);
            if (response.getStatusLine().getStatusCode() == 200) {
                HttpEntity resEntity = response.getEntity();
                String message = EntityUtils.toString(resEntity, "utf-8");
                /*
                            * {
                "errcode": 0,
                "errmsg": "ok",
                "access_token": "Ji6VwhG-CdB2J2Dch25xBeSIA6eG9adjatOGr22nDRfg2LiMsTLtYqHC1ikdXF8ROxsm1w8azrifPwsI52Fc73fIdSkmpOtE5D7jDL8WOB7WSlTsZzhMbHxNHRNAgbBDtXuRxjKfuqkdfEJSSRTFL5xrWVEVLrEfi_KmyufVjHItCWezF3gmbUgKUkFoWCdbI3p1oz66EaZk_Dv-GS1UZg",
                "expires_in": 7200
            }
                * */
                JSONObject jsonObject = JSONObject.parseObject(message);
                accessToken = jsonObject.getString("access_token");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return accessToken;
    }

    // 通过防疫小助手发送消息给USERID
    public static void pushtMsg(String accessToken, String requestStr) {

        HttpPost httppost = new HttpPost("https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=" + accessToken);
        // 设置报文和通讯格式
        StringEntity stringEntity = new StringEntity(requestStr, "utf-8");
        stringEntity.setContentEncoding("utf-8");
        stringEntity.setContentType("application/json");
        httppost.setEntity(stringEntity);
        try {
            HttpResponse response = httpclient.execute(httppost);
            if (response.getStatusLine().getStatusCode() == 200) {
                HttpEntity resEntity = response.getEntity();
                String message = EntityUtils.toString(resEntity, "utf-8");
                /*
                * {"errcode":0,"errmsg":"ok","type":"file","media_id":"3em8SWYPwKj5pizDxCsN5Et7fWig2Zk3jq20Qcb6PDraP0Gh6nJeZ1d0sgKDr9RkR","created_at":"1650682029"}
                * */
                JSONObject jsonObject = JSONObject.parseObject(message);
                int errcode = jsonObject.getIntValue("errcode");
                if(errcode == 0) {
                    System.out.println("===============发送消息成功===============");
                } else {
                    System.out.println("===============发送消息失败===============" + jsonObject.getString("errmsg"));
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 上传临时素材
     * 素材上传得到media_id，该media_id仅三天内有效
     * media_id在同一企业内应用之间可以共享
     * 返回空的
     * @param accessToken
     * @param file
     * @return
     */
    public static String fileUpload(String accessToken, File file) {
        String media_id = "";
        HttpPost httppost = new HttpPost("https://qyapi.weixin.qq.com/cgi-bin/media/upload?access_token=" + accessToken+"&type=file");
        FileBody filebody = new FileBody(file);
        HttpEntity reqEntity = MultipartEntityBuilder.create().addPart("media", filebody)
                .setMode(HttpMultipartMode.RFC6532)
                .build();
        httppost.setEntity(reqEntity);
        System.out.println("executing request " + httppost.getRequestLine());
        try {
            CloseableHttpResponse response = httpclient.execute(httppost);
            if (response.getStatusLine().getStatusCode() == 200) {
                HttpEntity resEntity = response.getEntity();
                String message = EntityUtils.toString(resEntity, "utf-8");
                JSONObject jsonObject = JSONObject.parseObject(message);
                int errcode = jsonObject.getIntValue("errcode");
                if (errcode == 0) {
                    System.out.println("===============发送临时文件成功===============");
                    media_id = jsonObject.getString("media_id");
                }else {
                    System.out.println("===============发送临时文件失败===============" + jsonObject.getString("errmsg"));
                }
                EntityUtils.consume(resEntity);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return media_id;
    }

    public static void main(String[] args) {
        String accessToken = getAccessToken();
        if(StringUtils.isEmpty(accessToken)) {
            System.out.println("获取企业微信accesstoken报错");
            return;
        }
        String requestStr = "{\"touser\":\""+USERID+"\",\"msgtype\":\"text\",\"agentid\":"+AGENTID+",\"text\":{\"content\":\"我要开始发送文件了！\"},\"safe\":0}";
        pushtMsg(accessToken, requestStr);
        File file = new File("C:\\Users\\99543\\Desktop\\tmp\\3居家隔离失联人员情况登记表.xlsx");
        String mediaId = fileUpload(accessToken, file);
        if(StringUtils.isEmpty(mediaId)) {
            System.out.println("上传临时素材失败");
            return;
        }
        String requestFileStr = "{\"touser\":\""+USERID+"\",\"msgtype\":\"file\",\"agentid\":"+AGENTID+",\"file\":{\"media_id\":\""+mediaId+"\"},\"safe\":0,\"enable_duplicate_check\":0,\"duplicate_check_interval\":1800}";
        pushtMsg(accessToken, requestFileStr);

    }
}
