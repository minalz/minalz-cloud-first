package cn.jorian.framework.jcloudprovider2.controller;
 
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
 
@RestController
@RequestMapping("/user")
public class UserController {
	@RequestMapping("/sayHello")
	public String sayhello() {
		return "I`m provider 2 ,Hello consumer sayHello!";
	}
	@RequestMapping("/sayHi")
	public String sayHi() {
		return "I`m provider 2 ,Hello consumer sayHi!";
	}
	@RequestMapping("/sayHaha")
	public String sayHaha() {
		return "I`m provider 2 ,Hello consumer sayHaha!";
	}
}