# SendMailRegularly

#### 读取excel内容，并以邮件的正文的方式定时发送
> 谁让老板总以发每日任务报告时间为加班时间呢，说好的弹性上班时间，尽然变成了只弹早上不弹晚上

- 示例: 设置时间段为0点50分-1点3分之间
```angular2html
starttime =   [[0,  50,  0]]     # 一个时间段的起始时间，hour, minute 和 second
endtime =   [[1, 3,  0]]    # 一个时间段的终止时间
```

- 执行: 
```angular2html
python main.py
```

