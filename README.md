# excel-eximport-wrapper

#### 介绍
{
excel-eximport-wrapper是针对现有数据进行的导入导出操作,现有数据可以是数据库中查询的数据，可以是代码运行中产生的临时数据，或者是外部excel文件中存在的数据，减少了使用者在开发过程中自行实现excel导入导出产生的代码量和开发时间的成本，通过本工具只需要告知导出或者导入的数据以及相关的个性化要求即可，个性化的定制包括某个条件下单元格的具体样式、表格间距等、个性化复杂的表格导出可以预先制作模板，根据模板将具体数据渲染。以上的种种具体的实现都由excel-eximport-wrapper内部实现}

#### 软件架构
本工具基于阿里的easyexcel进行开发，在此基础上做了一层封装，
演示工程基于spring boot搭建一个web项目


#### 安装教程

1.  本工具分为两个分支，分别为master和develop，可根据不同的需要进行分支代码的使用获取
2.  **master**：主要针对纯excel导入导出功能的开发，可以单独构建作为其他项目的依赖模块
3.  **develop**:主要是作为一个演示如何使用excel导入导出功能的开发，以一个web项目为例，在业务逻辑层对导入导出功能的调用，所以一般的开发流程是先在master上开发功能，然后在develop上合并master的代码，再进行测试或者演示使用

#### 使用说明

1.  导出功能</br>
1.1 EasyExcelUtil.writeToExcel(List<ExcelSheetElement> sheetList, OutputStream outputStream)
     </br>sheetList为excel中每一个sheet的属性，包括sheet上的数据、sheet名称、sheet上每个table个性化的样式、合并单元格等个性化设置，具体参照ExcelSheetElement属性,以及演示工程seivice层代码</br>
1.2 EasyExcelUtil.fillWithTemplate(String targetFilename,ExcelFillElement excelFillElement, OutputStream outputStream)
</br>excelFillElement为excel模板导出数据对，包括导出的数据象、单元格样式、合并单元格等个性化设置，OutputStream 为导出后excel文件存放位置，具体参照ExcelFillElement 属性,以及演示工程seivice层代码</br>
2.  导出功能


3.  xxxx

#### 参与贡献

1.  Fork 本仓库
2.  新建 Feat_xxx 分支
3.  提交代码
4.  新建 Pull Request


#### 特技

1.  使用 Readme\_XXX.md 来支持不同的语言，例如 Readme\_en.md, Readme\_zh.md
2.  Gitee 官方博客 [blog.gitee.com](https://blog.gitee.com)
3.  你可以 [https://gitee.com/explore](https://gitee.com/explore) 这个地址来了解 Gitee 上的优秀开源项目
4.  [GVP](https://gitee.com/gvp) 全称是 Gitee 最有价值开源项目，是综合评定出的优秀开源项目
5.  Gitee 官方提供的使用手册 [https://gitee.com/help](https://gitee.com/help)
6.  Gitee 封面人物是一档用来展示 Gitee 会员风采的栏目 [https://gitee.com/gitee-stars/](https://gitee.com/gitee-stars/)
