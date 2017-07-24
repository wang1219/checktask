## Required
Python >= 2.7

## Quick start
```py
 # 安装python虚拟环境命令virtualenv
 $ sudo yum install -y virtualenv

 # 进入项目路径
 $ cd path

 # 安装名为.venv的虚拟环境
 $ virtualenv .venv

 # 启用虚拟环境, 启用成功后虚拟环境名字会显示在目录最前面，可能像：(.venv) ➜  checktask git:(master) ✗
 $ source .venv/bin/activate

 # 安装依赖包
 $ pip install -r requirements.txt -i http://pypi.doubanio.com/simple/ --trusted-host pypi.doubanio.com

 # 运行
 $ python checktask.py

 # 运行完成后退出虚拟环境
 $ deactivate
```
