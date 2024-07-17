package com.ek.common.utils;

import java.io.BufferedReader;
import java.io.DataOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Map;
import java.util.Map.Entry;
import java.net.URLEncoder;
import java.util.HashMap;

public class NoteUtil {
		public static final String DEF_CHATSET = "UTF-8";
	    public static final int DEF_CONN_TIMEOUT = 30000;
	    public static final int DEF_READ_TIMEOUT = 30000;
	    public static String userAgent =  "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.66 Safari/537.36";
	    //配置您申请的KEY
	    public static final String APPKEY ="f3ad1de864cbe36fb3745a2a158d8626";
	    
	   
	    //发送短信
	    public static String sendNote(String phone,String type){
	        String result =null;
	        String key=(int)((Math.random()*9+1)*100000)+""; //6位数验证码
	        String url ="http://v.juhe.cn/sms/send";//请求接口地址
	        Map<String, Object> params = new HashMap<>();//请求参数
	            params.put("mobile",phone);//接收短信的手机号码
	            params.put("tpl_id",(type.equals("1")?"34312":"209791"));//短信模板ID，请参考个人中心短信模板设置
	            params.put("tpl_value","#code#="+key);//变量名和变量值对。如果你的变量名或者变量值中带有#&=中的任意一个特殊符号，请先分别进行urlencode编码后再传递，<a href="http://www.juhe.cn/news/index/id/50" target="_blank">详细说明></a>
	            params.put("key",APPKEY);//应用APPKEY(应用详细页查询)
	           // params.put("dtype","");//返回数据的格式,xml或json，默认json
	 
	        try {
	            result =net(url, params, "POST");
	            /*JSONObject object = JSONObject.fromObject(result);
	            if(object.getInt("error_code")==0){
	                System.out.println(object.get("result"));
	            }else{
	                System.out.println(object.get("error_code")+":"+object.get("reason"));
	            }*/
	            System.out.println(result);
	           
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	        return type.equals("1")?key:"17ca";
	    }
		private static String net(String strUrl, Map<String, Object> params, String method) throws Exception{
			 HttpURLConnection conn = null;
		        BufferedReader reader = null;
		        String rs = null;
		        try {
		            StringBuffer sb = new StringBuffer();
		            if(method==null || method.equals("GET")){
		                strUrl = strUrl+"?"+urlencode(params);
		            }
		            URL url = new URL(strUrl);
		            conn = (HttpURLConnection) url.openConnection();
		            if(method==null || method.equals("GET")){
		                conn.setRequestMethod("GET");
		            }else{
		                conn.setRequestMethod("POST");
		                conn.setDoOutput(true);
		            }
		            conn.setRequestProperty("User-agent", userAgent);
		            conn.setUseCaches(false);
		            conn.setConnectTimeout(DEF_CONN_TIMEOUT);
		            conn.setReadTimeout(DEF_READ_TIMEOUT);
		            conn.setInstanceFollowRedirects(false);
		            conn.connect();
		            if (params!= null && method.equals("POST")) {
		                try {
		                    DataOutputStream out = new DataOutputStream(conn.getOutputStream());
		                        out.writeBytes(urlencode(params));
		                } catch (Exception e) {
		                    // TODO: handle exception
		                }
		            }
		            InputStream is = conn.getInputStream();
		            reader = new BufferedReader(new InputStreamReader(is, DEF_CHATSET));
		            String strRead = null;
		            while ((strRead = reader.readLine()) != null) {
		                sb.append(strRead);
		            }
		            rs = sb.toString();
		        } catch (IOException e) {
		            e.printStackTrace();
		        } finally {
		            if (reader != null) {
		                reader.close();
		            }
		            if (conn != null) {
		                conn.disconnect();
		            }
		        }
		        return rs;
		}
		
		public static String urlencode(Map<String,Object>data) {
	        StringBuilder sb = new StringBuilder();
	        for (Entry<String, Object> i : data.entrySet()) {
	            try {
	                sb.append(i.getKey()).append("=").append(URLEncoder.encode(i.getValue()+"","UTF-8")).append("&");
	            } catch (UnsupportedEncodingException e) {
	                e.printStackTrace();
	            }
	        }
	        System.out.println(sb);
	        return sb.toString();
	    }
}
