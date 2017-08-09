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
$ cd /path/to/

# 安装虚拟环境及依赖包
$ python tools/install_venv.py

# 运行(Crontab时请使用绝对路径，如：/opt/app/tools/with_venv.sh python /opt/app/checktask.py)
$ tools/with_venv.sh python checktask.py

# Linux 配置Crontab定时任务(root账户)
$ crontab -e # 配置crontab，同VIM,写入下面命令，保存(注意更改为自己的目录)
0 10 * * * /opt/app/tools/with_venv.sh python /opt/app/checktask.py # 此命令代表每天早上10点执行，详情请自行百度，

$ crontab -l # 查看当前的定时任务
```

