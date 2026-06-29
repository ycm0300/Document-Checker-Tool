# document-compliance-checker

一个用于文档内容检查的小工具，目前处于初始开发阶段。

## 当前版本

V0.1.0

## 功能说明

当前版本主要用于读取 Word 文档内容，并进行基础检查验证。

## 目录说明

```text
document-compliance-checker/
├── Input/        # 输入文档目录，实际文档不上传
├── Output/       # 输出结果目录，实际结果不上传
├── read_word.py  # 主程序
├── README.md     # 项目说明
├── CHANGELOG.md  # 版本变更记录
└── .gitignore    # Git 忽略配置
```

## 使用说明

将需要检查的文档放入 `Input` 目录，然后运行脚本：

```bash
python read_word.py
```

检查结果输出到 `Output` 目录。

## 说明

`Input` 和 `Output` 目录中的实际文件不会提交到仓库，仅保留目录结构。
