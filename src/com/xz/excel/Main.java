package com.xz.excel;

import com.xz.excel.entity.Config;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.Properties;

/**
 * Created by Fingal on 2017-5-17.
 */
public class Main {
    public static void main(String[] args) {
        //1.初始化系统参数
        Properties prop = new Properties();
        try {
            prop.load(new InputStreamReader(new FileInputStream(new File("config/config.properties")),"UTF-8"));
        } catch (Exception e) {
            System.out.println("读取配置文件失败");
            return;
        }
        Config config = new Config();
        config.load(prop);
        System.out.println("需要统计文件:" + config.getSourceFile());
        System.out.println("统计结果 : " + config.getTargetFile());
        System.out.println(" ");

        // 2.执行
        System.out.println("开始");
        Statistics statistics = new Statistics();
        statistics.process(config);
        System.out.println("结束");
    }
}
