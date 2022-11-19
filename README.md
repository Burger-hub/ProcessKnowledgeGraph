# ProcessKnowlegeGraph
project for course

## Requirement

- Neo4j Desktop

- Anaconda

- python-occ

  - 在anaconda prompt中运行 `conda create -n pythonocct -c dlr-sc -c pythonocc pythonocc-core=7.4.0rc1`

  - anaconda装环境遇到无法定位程序输入点OPENSSL_sk_new_reserve……问题，查阅[anaconda装环境遇到无法定位程序输入点OPENSSL_sk_new_reserve……问题_只有小松了的博客-CSDN博客](https://blog.csdn.net/qq_37465638/article/details/100071259)
  - 输入`activate pythonocct`检验pythonocc环境是否安装成功

  - 设置默认python解释器为python-occ，`D:\Anaconda\envs\pythonocct\python.exe`

- py2neo
- openpyxl
- pyqt5

## Usage

- 将待用的step文件放入目录中
- 在pythonocct环境中运行 feature_extraction.py
- 点击某一特征作为待加工特征
- 输出待选加工路径
