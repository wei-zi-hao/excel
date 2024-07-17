package com.ek;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;

/**
 * @author EricWei on 2019/10/16
 */
@SpringBootApplication(exclude = {DataSourceAutoConfiguration.class})
public class EKApplication {
        public static void main(String[] args){
            SpringApplication.run(EKApplication.class,args);
        }
}
