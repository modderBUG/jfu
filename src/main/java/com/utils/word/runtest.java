package com.utils.word;

import com.jacob.activeX.ActiveXComponent;
/**
 * 测试类，查看环境是否配置正确。
 * **/
public class runtest {
    public static void main(String[] args) {
       	ActiveXComponent word = null;
       	try {
           	        word = new ActiveXComponent("Word.Application");
           	        System.out.println("jacob当前版本："+word.getBuildVersion());
                    System.out.println("jacob当前对象："+word.getObject());
           	}catch(Exception e ){
           	         e.printStackTrace();
           	}

    }
}
