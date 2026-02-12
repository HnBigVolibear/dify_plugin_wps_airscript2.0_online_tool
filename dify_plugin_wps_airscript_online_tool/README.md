## dify_plugin_wps_airscript2.0_online_tool

**Author:** HnBigVolibear 湖南大白熊工作室
**Version:** 2.0.1
**Type:** tool

### Description

![img](./_assets/wps.png)
一个Dify插件 —— 基于WPS AirScript2.0设计的Python客户端，配合配套的JS，轻松实现WPS在线智能表格的各类读写操作，支持各类日常操作40+种，开箱即用！效率拉满！现在已经封装成Dify插件，方便直接在工作流里使用！轻松打通Dify与WPS的数据交互！拒绝数据孤岛！

### Usage

1. use and see NewBee!
2. 建议直接使用我打包好的Dify插件安装包，直接安装即可使用！

### Quick Start Develop
- 注：python版本建议3.12及以上！
- 安装依赖：`pip install -r requirements.txt`
- 下载并解压本仓库，进入到当前`dify_plugin_wps_airscript_online_tool`文件夹
- 修改`.env`文件，填入你的Dify平台的插件调试密钥！
- 运行`python -m main.py`即可！回到Dify平台的插件页面，即可看到插件已经上线！


## **🛠️ 二、开始使用插件**
> 记住：**任何工作流操作前，都必须先完成初始化！**

本插件提供以下常用功能模块，每个节点都有详细说明，请仔细阅读：

- 📄 **单元格读写操作**
- 🎨 **格式化设置**（字体、颜色、对齐、边框等）
- 🔢 **行列操作**（插入、删除、调整大小）
- 🔍 **查找和替换**
- 📊 **排序和复制粘贴**
- 📑 **工作表管理**
- ⚡ **批量数据处理**

#### **首先进行初始化**
- 回到插件中的 **「初始化WPS_AirScript接口」** 节点
- 填入刚才获取的 `file_id`、`token`、`script_id`
- 点击运行，成功后会返回提示
- ✅ **小提示**：初始化成功后，可以关闭该节点的 **「是否返回帮助信息」** 参数

---

#### **接下来你就可以随意使用任意操作节点块块了！**
请自行探索使用，每个节点的每个参数都有详细说明！
**请务必仔细阅读参数说明！请务必仔细阅读参数说明！请务必仔细阅读参数说明！**
**阅读理解不行的同学，建议回小学重修语文哦～** 😉 

---

## **📚 其他说明**
#### 参考链接 & 鸣谢：
- [WPS 智能表格 API 文档](https://www.kdocs.cn/l/cftIrDJVIvCU)
- [应道社区讨论](https://www.yingdao.com/community/detaildiscuss?id=885400393968951296)
- **WPS 官方 AirScript 文档**：[点击查看](https://airsheet.wps.cn/docs/apitoken/intro.html)

---

### **👨‍💻 联系我**
- 湖南大白熊工作室 
- https://github.com/HnBigVolibear
如有技术问题或改进建议，欢迎联系：  
📧 **1486203070@qq.com**

---

> 让数据流动起来，让表格变得更聪明！祝你使用愉快！ 🚀

