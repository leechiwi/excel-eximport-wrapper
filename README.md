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

1.  导出功能  
**1.1 EasyExcelUtil.writeToExcel(List<ExcelSheetElement> sheetList, OutputStream outputStream)**  
sheetList为excel中每一个sheet的属性，包括sheet上的数据、sheet名称、sheet上每个table个性化的样式、合并单元格等个性化设置，具体参照ExcelSheetElement属性,以及演示工程seivice层代码  
**1.2 EasyExcelUtil.fillWithTemplate(String targetFilename,ExcelFillElement excelFillElement, OutputStream outputStream)**  
excelFillElement为excel模板导出数据对象，包括导出的数据、模板文件位置、单元格样式、合并单元格等个性化设置，其中单元格样式和合并单元格设置主要是针对带有循环扩展区域的新创建单元格以及再合并处理，否则新扩展的单元格是不会合并和添加任何样式的即使是模板中预先设置的单元格你已经做好了这些，targetFilename为导出的文件的名称，默认存放在和导出模板同目录下，OutputStream 为导出后excel文件存放位置，具体参照ExcelFillElement 属性,以及演示工程seivice层代码  
 <kbd>注意</kbd>:模板导出，需要满足模板的特定书写格式，可见阿里的模板导出格式说明  
**1.3 EasyExcelUtil.writeToExcelInZip(List<List<ExcelSheetElement>> excelList,List<String> excelFileNames, OutputStream zipOutputStream)**  
针对导出多个excel文件,则可以将多个文件导出后统一放到压缩包中，代码层面是 1.1的增强，excelList中每一个元素就是一个excel文件对应的属性，等同于 **1.1** 的sheetList，excelFileNames为对应的excel文件名称，如果没有此参数 ，则默认以excelList元素下表数字作为每一个excel文件名称 ,zipOutputStream为zip文件的输出流
2.  导入功能  
**2.1 EasyExcelUtil.List<ExcelSheetElement> getFromExcel(InputStream inputStream, List<ExcelSheetElement> sheetElements, List<String> includeSheets, Consumer<?> defaultConsumer,List<String> defaultHead)**  
inputStream为导入的excel文件,sheetElements为对应的导入文件的每一个sheet对象，包括每一个sheet数据对应的接收实体类型和具体叫收到数据后的操作，比如插入数据库操作等，includeSheets用来控制文件中想导入的sheet即用户可以选择导入哪个sheet，对于多个sheet可能会存在对应的是同一个实体类的情况,所以就不需要重复设置每一个sheet对应的属性而是通过defaultConsumer, defaultHead参数统一设置即可。返回值List<ExcelSheetElement>可以获取到具体的每个sheet对应的导入数据

3.  web导出功能  
**3.1 EasyExcelUtil.exportWebExcel(List<ExcelSheetElement> sheetList, HttpServletRequest request, HttpServletResponse response, String filename)**  
该功能是针对web项目的excel导出功能，代码层面是**1.1**方法的一个增强,通过request和response来构造输出流即**1.1**方法的outputstream参数  
**3.2 EasyExcelUtil.exportWebExcelInZip(List<List<ExcelSheetElement>> excelList,List<String> excelFileNames, HttpServletResponse response, String filename)**  
该功能是针对web项目的多个excel导出并打包成zip的功能，代码层面是**1.3**方法的一个增强,通过response来构造输出流即**1.3**方法的zip的outputstream参数  


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
