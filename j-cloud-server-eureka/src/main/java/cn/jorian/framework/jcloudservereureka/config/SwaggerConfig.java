package cn.jorian.framework.jcloudservereureka.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import springfox.documentation.builders.ApiInfoBuilder;
import springfox.documentation.builders.PathSelectors;
import springfox.documentation.builders.RequestHandlerSelectors;
import springfox.documentation.service.ApiInfo;
import springfox.documentation.spi.DocumentationType;
import springfox.documentation.spring.web.plugins.Docket;
import springfox.documentation.swagger2.annotations.EnableSwagger2;

/**
 * swagger2的配置类
 */
@Configuration
@EnableSwagger2
public class SwaggerConfig {

   @Bean
   public Docket createRestApi() {
      return new Docket(DocumentationType.SWAGGER_2).apiInfo(apiInfo()).select()
            // api扫包范围
            .apis(RequestHandlerSelectors.basePackage("cn.jorian.framework.jcloudservereureka.controller")).paths(PathSelectors.any()).build();
   }

    /**
     * 创建该API的基本信息（这些基本信息会展现在文档页面中）
     * 访问地址：http://项目实际地址/swagger-ui.html
     * @return
     */
   private ApiInfo apiInfo() {
      return new ApiInfoBuilder().title("Swagger接口发布测试").description("测试|Swagger接口功能")
            .termsOfServiceUrl("http://www.baidu.com")
            .version("1.0").build();
   }

}