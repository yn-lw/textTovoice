package com.hbsi.demo;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.*;

/**
 * @author JLW
 * @create 2021-05-15 18:52
 */
public class Demo01 {

    public static void main(String[] args) {
        //传入一个txt文本
        readText(new File("D:/测试.txt"));
    }

    public static void readText(File text) {
        Reader reader=null;
        BufferedReader br=null;
        ActiveXComponent sap = new ActiveXComponent("Sapi.SpVoice");
        Dispatch sapo=sap.getObject();
        try {
            reader=new FileReader(text);
            br =new BufferedReader(reader);
            //音量
            sap.setProperty("Volume", new Variant(100));
            //语速
            sap.setProperty("Rate",new Variant(0.1));
            String info="";
            while((info=br.readLine())!=null) {
                //语音播放当前行的内容
                Dispatch.call(sapo,"Speak",new Object[]{info});
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            sapo.safeRelease();
            sap.safeRelease();
            try {
                reader.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
