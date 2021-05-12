package com.ruoyi.bigdataAnalysis.until;

import com.ruoyi.common.core.utils.StringUtils;

import javax.net.ssl.*;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;

public class UrlFileUtil {

    public static File getNetUrl(String netUrl) {
        //判断http和https
        File file = null;
        if (netUrl.startsWith("https://")) {
            file = getNetUrlHttps(netUrl);
        } else {
            file = getNetUrlHttp(netUrl);
        }
        return file;
    }


    /**
     * 获取http 文件
     *
     * @param netUrl
     * @return
     */
    public static File getNetUrlHttp(String netUrl) {
        //对本地文件命名
        String fileName = reloadFile(netUrl);
        File file = null;


        URL urlfile;
        InputStream inStream = null;
        OutputStream os = null;
        try {
            file = File.createTempFile("net_url", fileName);
            //下载
            urlfile = new URL(netUrl);
            inStream = urlfile.openStream();
            os = new FileOutputStream(file);

            int bytesRead = 0;
            byte[] buffer = new byte[8192];
            while ((bytesRead = inStream.read(buffer, 0, 8192)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (null != os) {
                    os.close();
                }
                if (null != inStream) {
                    inStream.close();
                }

            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        return file;
    }

    /**
     * 获取到本地(支持https)
     *
     * @param fileUrl 远程地址
     * @throws Exception
     */
    public static File getNetUrlHttps(String fileUrl) {
        //对本地文件进行命名
        String file_name = reloadFile(fileUrl);
        File file = null;

        DataInputStream in = null;
        DataOutputStream out = null;
        try {
            file = File.createTempFile("net_url", file_name);

            SSLContext sslcontext = SSLContext.getInstance("SSL", "SunJSSE");
            sslcontext.init(null, new TrustManager[]{new X509TrustUtiil()}, new java.security.SecureRandom());
            URL url = new URL(fileUrl);
            HostnameVerifier ignoreHostnameVerifier = new HostnameVerifier() {
                @Override
                public boolean verify(String s, SSLSession sslsession) {
                    return true;
                }
            };
            HttpsURLConnection.setDefaultHostnameVerifier(ignoreHostnameVerifier);
            HttpsURLConnection.setDefaultSSLSocketFactory(sslcontext.getSocketFactory());
            HttpsURLConnection urlCon = (HttpsURLConnection) url.openConnection();
            urlCon.setConnectTimeout(6000);
            urlCon.setReadTimeout(6000);
            int code = urlCon.getResponseCode();
            if (code != HttpURLConnection.HTTP_OK) {
                throw new Exception("文件读取失败");
            }
            // 读文件流
            in = new DataInputStream(urlCon.getInputStream());
            out = new DataOutputStream(new FileOutputStream(file));
            byte[] buffer = new byte[2048];
            int count = 0;
            while ((count = in.read(buffer)) > 0) {
                out.write(buffer, 0, count);
            }
            out.close();
            in.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (null != out) {
                    out.close();
                }
                if (null != in) {
                    in.close();
                }
            } catch (Exception e2) {
                e2.printStackTrace();
            }
        }

        return file;
    }

    /**
     * 重命名，UUIU
     *
     * @param oleFileName
     * @return
     */
    public static String reloadFile(String oleFileName) {
        oleFileName = getFileName(oleFileName);
        if (StringUtils.isEmpty(oleFileName)) {
            return oleFileName;
        }
        //得到后缀
        if (oleFileName.indexOf(".") == -1) {
            //对于没有后缀的文件，直接返回重命名
            return UUID.randomUUID().toString().replace("-", "");
        }
        String[] arr = oleFileName.split("\\.");
        // 根据uuid重命名图片
        String fileName = UUID.randomUUID().toString().replace("-", "") + "." + arr[arr.length - 1];

        return fileName;
    }

    /**
     * 把带路径的文件地址解析为真实文件名
     *
     * @param url
     */
    public static String getFileName(final String url) {
        if (StringUtils.isEmpty(url)) {
            return url;
        }
        String newUrl = url;
        newUrl = newUrl.split("[?]")[0];
        String[] bb = newUrl.split("/");
        String fileName = bb[bb.length - 1];
        return fileName;
    }


}
