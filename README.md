## 关于 Mail Editor

本项目已在[Github](http://github.com/Tresol/MailEditor)开放源代码，遵循[GNU通用公共许可证](https://www.gnu.org/licenses/gpl-3.0.html#license-text)协议（参阅LICENSE文件）。

[Tresol](http://github.com/Tresol)为项目负责人，电子邮件为[tresol@163.com](tresol@163.com)。

英文翻译仅供参考。一切最终以中文文本为准。[英文文本 English Version](https://github.com/Tresol/MailEditor/blob/main/README.en.md)

### 最新版本（1.0.1）更新说明
最新版本为1.0.1版本，已于2023年6月30日更新。新版特性如下：

1. 邮件页面启用基于Microsoft 365 通知邮件的HTML格式，版式更加美观；
2. 文件选项中增加了“仅保存”选项；
3. 优化代码，优化附属文件。

更新日志，请参阅 [Update.md](https://github.com/Tresol/MailEditor/blob/main/Update.md)。

### 功能说明

本项目适合将文件批量传送至打印室。

### 运行环境

由于本项目尚未使用GUI渲染，故用户需要安装[Python](http://python.org/)（3.7+）。

该项目所有内容均在Windows10，Python 3.11环境下测试。

### 使用说明

使用前，先编辑配置文件`mail.ini`，填写邮箱SMTP服务器地址、发件人的用户名和密码（部分邮件服务商为授权码）、收件人；

然后，将需要发送的文件复制到`inbox`文件夹，执行`MailEditor.py`文件并按提示操作；

用户所有的发送存档以及日志将保存在`archives`文件夹。

### 鸣谢

[Github](http://github.com/)为项目发布提供了良好的平台生态；

[Microsoft 365](http://aka.ms)为项目提供了具有参考价值的邮件模板。
