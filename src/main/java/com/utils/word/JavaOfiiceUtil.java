package com.utils.word;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;


public class JavaOfiiceUtil {

    public static void main(String[] args) {
        String inputFIle ="C:\\Users\\modderBUG\\Desktop\\java服务端使用jcob操作office.doc";
        String outputFile ="C:\\Users\\modderBUG\\Desktop\\out.doc";
        extractDoc(inputFIle,outputFile);
    }

    public static void extractDoc(String inputFIle, String outputFile) {
       boolean flag = false;
       // 打开Word应用程序
       ActiveXComponent app = new ActiveXComponent("Word.Application");
        try {
           // 设置word不可见
           app.setProperty("Visible", new Variant(false));
           // 打开word文件
           Dispatch doc1 = app.getProperty("Documents").toDispatch();
           Dispatch doc2 = Dispatch.invoke(
                 doc1,
                 "Open",
                 Dispatch.Method,
                 new Object[] { inputFIle, new Variant(false),
                       new Variant(true) }, new int[1]).toDispatch();
           // 作为txt格式保存到临时文件
           Dispatch.invoke(doc2, "SaveAs", Dispatch.Method, new Object[] {
                 outputFile, new Variant(7) }, new int[1]);
           // 关闭word
           Variant f = new Variant(false);
           Dispatch.call(doc2, "Close", f);
           flag = true;
        } catch (Exception e) {
           e.printStackTrace();
        } finally {
           app.invoke("Quit", new Variant[] {});
        }
        if (flag == true) {
           System.out.println("Transformed Successfully");
        } else {
           System.out.println("Transform Failed");
        }
     }

}
