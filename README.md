# 基于nodejs的excel表格处理插件，所有想要运行需要先安装nodejs环境

### 背景

PMO表格统计中，各个项目分不同sheet收集数据，但最终数据需要汇总和统计

### 使用方法一

1. 把需要合并的文件放到excel文件夹里，如果excel中有多个文件，需要在 index.js 里 init 方法中指定需要处理的文件名
2. 第一次执行需要运行一次init.bat文件，以后每次只需运行start.bat文件
3. 到result下拿到合并完成的excel

### 使用方法二

```bash
$ npm i
$ npm run start
```

#### 注意

excel文件里各sheet的项目模块功能要保持4列，各表尽量保持格式统一
