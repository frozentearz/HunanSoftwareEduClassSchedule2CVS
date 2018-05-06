## 实现让 Siri 响应用户的：“嘿, Siri. 我在 xx 时间有什么课?”

### Installation
1.   下载项目
```shell
# git clone https://github.com/frozentearz/HunanSoftwareEduClassSchedule2CVS.git
```
2.   进入项目
```shell
# cd HunanSoftwareEduClassSchedule2CVS
```
3.   安装依赖库
```shell
Python2： #pip install -r requirements.txt

Python3： #pip3 install -r requirements.txt
```

### Usage:
1. 把课程表放到 *xlsx* 文件夹下，并命名为ClassSchedule.xlsx  
`PS: 注意此文件夹下同时只能放一个 xlsx 文件，且目前仅支持 xlsx 文件`

2. 执行，得到 「课表.cvs」
```Python
# python3 RunMe.py
```

3. 打开 [Google 日历](https://calendar.google.com)，并登陆你的 Google 账号。    
`PS: 打开链接需要科学上网(即FQ)`

4. 点击右上角的设置图标 -> 设置 -> 时区，把时区设置为你所在的时区。  
目前仅支持： `(GMT+08:00) 中国时间 - 北京`

5. 点击右上角的设置图标 -> 设置 -> 基本 -> 添加日历 -> 新建日历，  
然后填写名称，说明，时区与所有者，再点击右下的创建日历。  
Eg: 
    ```
    名称：大二课表
    说明：湖南软件学院Java17届下学期大二课表
    时区：选择 (GMT+08:00) 中国时间 - 北京
    所有者：frozentearzg@gmail.com
    ```

6. 点击右上角的设置图标 -> 设置 -> 基本 -> 导入和导出 -> 导入，  
选择从计算机选择文件，选中 RunMe.py 生成的 课表.cvs 文件，  
添加至日历，选择刚刚创建的 「大二课表」 ，点击导入，完成。

7. 这个时候你已经可以在 [Google 日历](https://calendar.google.com) 看到你的日常课程安排了，下一步把 Google 日程安排导入到 Apple 的日历，让 Siri 可以访问

8. 详情见 Google Support: (https://support.google.com/calendar/answer/99358?hl=zh-Hans)  
如果你不能科学上网，憋方，我帮你复制好了: 

 - 在 Apple 日历上查看 Google 日历活动
    - 您计算机上的 Apple 日历
>**注意**：请确保您的计算机上安装的是最新版本的 Apple 日历和 Apple 操作系统。了解如何[查找您的 Apple 计算机操作系统](https://support.apple.com/HT201260)或[更新您的计算机操作系统](https://support.apple.com/HT201541)。  
在计算机上打开日历 Apple 日历![Apple 日历](./images/AppleCalendar.png)。  
在屏幕左上角，点击**日历 > 偏好设置**。  
点击**帐号**标签。  
点击“帐号”标签左侧的“添加”图标 ![添加](./images/add.png)。  
选择 **Google** > **继续**。  
要添加您的 Google 帐号信息，请按照屏幕上的步骤操作。  
在“帐号”标签中，使用“刷新日历”即可选择要让 Apple 日历与 Google 日历同步的频率。  
    - 您 iPhone 或 iPad 上的日历应用
>您可以将 Google 日历与 iPhone 或 iPad 自带的日历应用同步。  
打开 iPhone 或 iPad 上的“设置”应用。  
点按**日历**。  
点按**帐号**。  
点按**添加帐号** > **Google**。  
输入您的电子邮件地址 接着 **下一步**。  
输入您的密码。如果您的操作系统不是最新版本且您使用[两步验证](https://support.google.com/accounts/answer/180744)，请输入[应用专用密码](https://security.google.com/settings/security/apppasswords)，而非普通密码。  
触摸**下一步**。  
系统会立即将邮件、通讯录和日历活动直接与您的 Google   帐号同步。要仅同步日历，请关闭其他服务。  
在 iPhone 上打开日历应用，以查看 Google 日历活动。  

**Ending**:   
之后你就能看到你的 Apple 日历上有你的课程安排，接下来就可以通过 Siri 询问有关课程安排，**Finished**.

### Todolist:
- 单双周课程分离
- 支持用户输入开学日期
- xls 与 xlsx 兼容

### Any Question, new a issues~