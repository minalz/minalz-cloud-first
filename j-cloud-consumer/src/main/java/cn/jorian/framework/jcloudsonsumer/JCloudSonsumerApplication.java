package cn.jorian.framework.jcloudsonsumer;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.cloud.netflix.eureka.EnableEurekaClient;

@SpringBootApplication
@EnableEurekaClient // 表明这是一个客户端
public class JCloudSonsumerApplication {

    public static void main(String[] args) {
        SpringApplication.run(JCloudSonsumerApplication.class, args);
    }

}
