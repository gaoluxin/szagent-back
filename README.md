# 收资工具-后台

将客户版本的收资表转换成研发需要的收资表。

## 功能模块

- 储能收资
  - 信息收资：将客户版储能收资表转换为研发版储能收资表
  - 点表收资：暂未开发
- 光伏收资：暂未开发

## 安装

```bash
pip install -r requirements.txt
```

## 使用方式

### 1. Web API服务

启动服务：

```bash
python run.py server
```

指定端口：

```bash
python run.py server --port 8080
```

开发模式（热重载）：

```bash
python run.py server --reload
```

访问API文档：http://localhost:8000/docs

### 2. 命令行工具

储能收资：

```bash
python run.py cli energy-storage 客户收资表.xlsx
```

指定输出文件：

```bash
python run.py cli energy-storage 客户收资表.xlsx -o 输出文件.xlsx
```

光伏收资（暂未开发）：

```bash
python run.py cli pv 客户收资表.xlsx
```

## 项目结构

```
szagent-back/
├── app/
│   ├── api/                 # API路由
│   │   └── routers/
│   │       ├── energy_storage.py
│   │       └── pv.py
│   ├── cli/                 # 命令行工具
│   │   └── main.py
│   ├── core/                # 核心配置
│   │   └── config.py
│   ├── models/              # 数据模型
│   │   └── schemas.py
│   └── services/            # 业务逻辑
│       ├── excel_reader.py
│       └── excel_writer.py
├── public/                  # 公共文件
│   ├── 储模板/
│   │   ├── 储能收资模板.xlsx
│   │   └── 储能点表模板.xlsx
│   └── 储能电站收资表-客户版.xlsx
├── output/                  # 输出目录
├── main.py                  # FastAPI应用入口
├── run.py                   # 模块总入口
└── requirements.txt         # 依赖包
```

## API接口

### 储能收资

- POST `/api/energy-storage/info-collection` - 储能信息收资
- GET `/api/energy-storage/health` - 健康检查

### 光伏收资

- GET `/api/pv/health` - 健康检查

## 技术栈

- Python 3.8+
- FastAPI
- openpyxl
- pandas
