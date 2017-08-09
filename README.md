## Required
Python >= 2.7

## Quick start
```sh
 # 安装python虚拟环境命令virtualenv
 $ sudo yum install -y virtualenv

 # 进入项目路径
 $ cd path

 # Tip
 使用前请先配置邮箱相关信息

 ## Quick start
```sh
$ cd /path/to/

# 安装虚拟环境
$ python tools/install_venv.py
# 安装依赖包
$ tools/with_venv.sh python setup.py install

# 运行(Crontab时请使用绝对路径，如：/opt/app/tools/with_venv.sh python /opt/app/checktask.py)
$ tools/with_venv.sh python checktask.py
```

