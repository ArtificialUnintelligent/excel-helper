### excel导出工具

#### 依赖
```xml
<dependency>
    <groupId>com.au</groupId>
    <artifactId>excel-helper</artifactId>
    <version>1.0-SNAPSHOT</version>
</dependency>
```

#### 接入说明
0. 提供一个导出相关的注解及对应注解解析器
0. 具体导出格式见result.xls，使用demo见单测类ExcelWriterTest
1. 目前只支持2级目录的情况，一级目录支持合并单元格，二级目录支持单个单元格，由于excel类型多样，导出时也较为复杂，难以统一，因此建议其他情况自行编写实现
2. 在实际使用poi包提供的excel导出相关工具时，数据获取和格式渲染都有可能成为性能瓶颈，导致导出速度过慢，尤其是对每个单元格格式的渲染，会严重拖慢生成文件的速度
3. 如果使用复杂的excel模板，固定部分可以在web服务容器启动时直接加载模板数据到内存，或缓存中间件，这样也会减少导出所消耗的时间
4. web服务做excel导出（实时数据）通常使用异步的形式，一个接口导出+上传到备份服务器，并写入进度到内存（通常使用Threadlocal将数据保存到线程，或者写入非关系型数据库，这里也更推荐非关系型数据库），另一个接口提供心跳实时同步前端导出进度
5. 低版本的poi或可能导致oom，具体参照https://github.com/OpenRefine/OpenRefine/issues/940 和 https://stackoverflow.com/questions/33368612/gc-overhead-limit-exceeded-with-apache-poi