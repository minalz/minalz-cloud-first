package cn.minalz.cloud.jcloudconsumerfeign.feign;

import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

@FeignClient("PROVIDER-USER")
public interface UserClient {

    @RequestMapping(value = "/user/sayHello", method = RequestMethod.GET)
    public String sayHello();

    @RequestMapping(value = "/user/sayHi", method = RequestMethod.GET)
    public String sayHi();

    @RequestMapping(value = "/user/sayHaha", method = RequestMethod.GET)
    public String sayHaha();
}
