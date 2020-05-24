package cn.jorian.framework.jcloudservereureka.controller;

import cn.jorian.framework.jcloudservereureka.entity.UserEntity;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiImplicitParam;
import io.swagger.annotations.ApiImplicitParams;
import io.swagger.annotations.ApiOperation;
import javafx.application.Application;
import org.springframework.web.bind.annotation.*;
@Api(value = "会员接口")
@RestController
public class DemoController {

    @ApiOperation(value = "swagger接口测试demo", nickname = "swagger接口测试demo昵称")
    @GetMapping("/getDemo")
    public String getDemo() {
        return "getDemo方法调用成功...";
    }

    @ApiOperation(value = "获取会员信息接口", nickname = "根据userName获取用户相关信息")
    @ApiImplicitParam(name = "userName", value = "用户名称", required = true, dataType = "String")
    @PostMapping("/postMember")
    public String postMember(@RequestParam String userName) {
        return userName; 
    }


    @ApiOperation(value = "添加用户信息", nickname = "nickname是什么", notes = "notes是什么", produces = "application/json")
    @PostMapping("/postUser")
    @ResponseBody
    @ApiImplicitParam(paramType = "query", name = "userId", value = "用户id", required = true, dataType = "int")
    public UserEntity postUser(@RequestBody UserEntity user, @RequestParam("userId") Integer userId) { // 这里用包装类竟然报错
        if (user.getId() == userId) {
            return user;
        }
        return new UserEntity();
    }


    @ApiOperation(value = "添加用户信息", nickname = "哈哈测试", notes = "哈哈测试添加用户", produces = "application/json") 
    @PostMapping("/addUser")
    @ResponseBody
    @ApiImplicitParams({
            @ApiImplicitParam(paramType = "query", name = "userName", value = "用户姓名", required = true, dataType = "String"), 
            @ApiImplicitParam(paramType = "query", name = "id", value = "用户id", required = true, dataType = "int") })
    public UserEntity addUser(String userName, Integer id) {
        UserEntity userEntity = new UserEntity();
        userEntity.setName(userName);
        userEntity.setId(id);
        return userEntity;
    }

    @ApiOperation(value = "获取第一条信息", nickname = "第一条信息", notes = "测试获取第一条信息是否成功", produces = "application/json")
    @ApiImplicitParams({
            @ApiImplicitParam(paramType = "query", name = "message", value = "信息", required = true, dataType = "String"),
            @ApiImplicitParam(paramType = "query", name = "userId", value = "用户🆔", required = true, dataType = "int")
    })
    @PostMapping("/firstInfo")
    @ResponseBody
    public String getFirstInfo(String message, Integer userId){
        return "获取第一条信息成功了！！" + userId + "收到了" + message;
    }

}