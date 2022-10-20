# WeichatBackupToDocx

### 目的

本项目主要是为了将聊天记录由微信（安卓）转出到docx文件中，以便长久的留存。

### 使用方法

1. 利用网上的教程从手机导出微信数据库。
   - 本人手机为华为，相应操作连接如下：https://www.louyue.com/huaweiwechat.htm。
   - 其余品牌手机也类似，或可ROOT。
2. 依次找到avatar（头像）、Download（下载文件）、emoji（表情包）、image2（图片）、video（视频）、voice2（语音）这几个文件夹以及system_config_prefs.xml这个文件
   - 教程：https://greycode.top/posts/android-wechat-bak/
3. 运行fmd_wechatdecipher.py脚本解密加密数据库。
   - 拨号界面输入*#06#，将imei替换为自己的imei
4. 调用wcdb中函数使用所需功能