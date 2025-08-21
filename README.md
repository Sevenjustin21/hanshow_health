# hanshow_health

定时抓取 Hanshow 海外系统健康监控数据，并生成每日小时级文字汇总表。

---

##  项目结构

```
hanshow_health/
├── main.py                  # 主程序入口
├── modules/                 # 模块化代码
│   ├── config.py            # 配置项与品牌映射
│   ├── login_handler.py     # 模拟登录并抓取表格数据
│   ├── extractor.py         # 提取数据、格式化文字并写入结果文件
│   └── utils.py             # 时间判断与重复写入判断工具
├── output/                  # 抓取与写入的 Excel 结果目录（自动生成）
├── .gitignore
├── requirements.txt
└── README.md
```

---

## ️ 功能说明

- 自动登录系统健康平台，抓取每小时设备状态数据；
- 汇总生成固定格式文字描述，按日期列写入 `hourly_summary_text.xlsx`；
- 支持调试模式，可跳过时间判断与写入去重；
- 输出文件统一保存在 `output/` 目录。

---

##  使用方法

1. 安装依赖：

```bash
pip install -r requirements.txt
```

2. 修改配置（如有需要）：
   - 在 `modules/config.py` 中填写ChromeDriver 路径等参数
   - 自己创建env文件填入用户名、密码信息

3. 运行脚本：

```bash
python main.py
```

---

##  调试模式（可选）

若需跳过整点判断和重复写入判断，在 `config.py` 中将 `DEBUG_MODE` 设置为 `True`：

```python
DEBUG_MODE = True
```

---

##  输出文件说明

- `health_table.xlsx`：抓取得到的原始数据表
- `hourly_summary_text.xlsx`：每小时生成的汇总描述文字，按日期横向排列
