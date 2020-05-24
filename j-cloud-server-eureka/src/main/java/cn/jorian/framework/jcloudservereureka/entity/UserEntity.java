package cn.jorian.framework.jcloudservereureka.entity;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;

/**
 * 创建时间: 23:09 2018/9/19
 * 修改时间:
 * 编码人员: ZhengQf
 * 版   本: 0.0.1
 * 功能描述:
 */
@ApiModel(value = "用户模型")
public class UserEntity {
    @ApiModelProperty(value="id" ,required= true,example = "123")
    private Integer id;
    @ApiModelProperty(value="用户姓名" ,required=true,example = "郑钦锋")
    private String name;


    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @Override
    public String toString() {
        return "DemoDoctor [id=" + id + ", name=" + name + "]";
    }

}