# python-projects
## pip访问太慢问题,使用清华大学镜像
### linux 
    linux下，修改 ~/.pip/pip.conf (没有就创建一个)， 修改 index-url至tuna，内容如下：
     [global]
    index-url = https://pypi.tuna.tsinghua.edu.cn/simple
### windows
    windows下，直接在user目录中创建一个pip目录，如：C:\Users\xx\pip，新建文件pip.ini，内容如下
    [global]
    index-url = https://pypi.tuna.tsinghua.edu.cn/simple


