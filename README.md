# SendMailRegularly

#### 读取excel内容，并以邮件的正文的方式定时发送
> 每天下班定时发送日报，防止忘记了

- 示例: 
设置时间段为16点50分-18点30分之间，会随机取出一个区间的时间值，作为发送时间
```angular2html
starttime =   [16,  50,  0]     # 起始时间，hour, minute 和 second
endtime =   [18, 30,  0]    # 终止时间, hour, minute 和 second
```

- 执行: 
```angular2html
python main.py
```

### 问题
- 运行报错: ImportError: Install xlrd >= 0.9.0 for Excel support错误的解决
```
在执行pandas读取excel的操作时,出现了问题,代码如下:

data = pd.read_excel(discfile)
 

ImportError: Install xlrd >= 0.9.0 for Excel support

解决办法：

需要pip安装xlrd的库,并且在当前代码中import这个xlrd这个库
```

