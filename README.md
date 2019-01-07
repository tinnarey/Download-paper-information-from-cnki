# Python-
爬取知网页面的文献信息，并存在Excel的内

环境：Python 3，谷歌浏览器，相应的包（pip install **）

需要手动更改的：

1、网页URL 也就是代码里面的web_url。这个是在你需要爬取的页面———F12 或者右键 检查 ———出现页面的network ———name处 
有个开头是brief.aspx? 双击——复制双击出来页面的网址 ，在替换代码‘ ’这个里面

2、把1 复制的网址里面curpage= 后面的数字改为 %s

3、cookie  1里面有个开头是brief.aspx? 这里 点击一下 右侧有个cookie 复制 替换代码的cookie

4、cookie 里面的时间需要改一下 可以爬的久一下  修改规则代码有说明
