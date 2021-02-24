# 说明

本脚本用于检测给定目录中所有试卷(.doc,.docx)中的重复题目。

定已**重复**为：给定两道题，若题干中的中文部分[编辑距离](https://baike.baidu.com/item/%E7%BC%96%E8%BE%91%E8%B7%9D%E7%A6%BB#:~:text=%E7%BC%96%E8%BE%91%E8%B7%9D%E7%A6%BB%E6%98%AF%E9%92%88%E5%AF%B9%E4%BA%8C,%E6%98%AF%E6%AF%94%E8%BE%83%E5%8F%AF%E8%83%BD%E7%9A%84%E5%AD%97%E3%80%82)小于等于1，则为重复。

- 可修改脚本中的常量`DISTANCE_THRESHOLD=1`来调整该阈值

## 使用环境需求

- 操作系统：Windows 7/10
- 需要软件：Microsoft Office Word 2010 或以上版本
- Python3
  - 下载并安装[Python3](https://www.python.org/download/releases/3.0/)，安装完毕后使用如下命令检查版本，确保输出版本值3.x.x

  ```bash
  python --version
  ```

  - 安装脚本依赖包

  ```bash
  pip install pywin32 editdistance
  ```

## 使用说明

- 将需要解析的全部试卷置于一个目录下（可以存在子目录），如`"C:\Papers"`
- 将脚本`duplicate_check.py`下载至某处，如`"C:\"`
- 在Windows自带命令行工具中运行该脚本

```bash
python duplicate_check.py -i C:\Papers -o '<分析结果存放文件夹>'
```

注：将`<分析结果存放文件夹>`替换为真实路径，如`"C:\Results"`

运行完毕后，将在指定`<分析结果存放文件夹>`下生成若干分析文件，包括：

- 一个`汇总结果.txt`：用于存放所有试卷的分析结果
- 若干个`[分析结果] <试卷名>.txt`：用于存放单个试卷的分析结果
