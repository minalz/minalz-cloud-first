package cn.jorian.framework.jcloudprovider2;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.cloud.netflix.eureka.EnableEurekaClient;

@SpringBootApplication
@EnableEurekaClient // 表明这是一个客户端
public class JCloudProvider2Application {

    public static void main(String[] args) {
        SpringApplication.run(JCloudProvider2Application.class, args);
    }

}
