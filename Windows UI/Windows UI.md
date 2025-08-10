# Winodws UI TA

开源项目：Python-UIAutomation-for-Windows：

https://github.com/yinkaisheng/Python-UIAutomation-for-Windows/blob/master/readme_cn.md


## 封装
在实际项目使用中发现 UIAutomation 的正确率比较的低，UIAutomation 的树遍历通常结合深度优先搜索（DFS）和广度优先搜索（BFS）。

在操作企业级软件的场景下，由于企业级软件本身比较复杂，组件加载和变换其实存在很多延迟，所以导致在搜索组件的过程中容易报错。

我对UIAutomation项目中提供的原始方法，进行一些封装，增加了一些重试机制，在实际场景下极大了增加了正确率

