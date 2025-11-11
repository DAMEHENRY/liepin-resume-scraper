# Automation Setup

A brief description of your project. What does it do? Who is it for?

## 介绍

本项目是一个基于 Python 和 Playwright 的猎聘网（Liepin）自动化工具，旨在高效筛选和获取符合特定需求的候选人信息。

核心功能包括：
1.  **自动化搜索与筛选**：根据用户输入的目标公司和职位，在猎聘网上自动执行搜索。
2.  **智能简历匹配**：集成火山引擎（VolcEngine）大语言模型，根据用户定义的访谈提纲，智能判断简历与岗位的匹配度。
3.  **多维度过滤**：支持根据候选人的最晚离职日期等条件进行初步筛选。
4.  **联系方式获取**：自动化模拟点击操作，以获取候选人的联系方式（云电话），并能处理图片格式的电话号码（通过截图保存）。
5.  **数据导出**：将所有符合条件的候选人信息（包括姓名、职位、公司、在职时间、联系方式和简历链接）整理并保存到 Excel 文件中。
6.  **交互式控制**：支持在运行过程中使用 `ESC` 键暂停/继续任务，并可在一次运行结束后选择是否开始新的搜索。

## 安装

如何安装和设置您的项目。例如：

```bash
# 1. 克隆仓库
git clone https://github.com/your-username/your-repository-name.git
cd your-repository-name

# 2. (可选) 创建并激活虚拟环境
python -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`

# 3. 安装依赖
pip install -r requirements.txt
```

## 使用

如何使用此脚本。

```bash
python main_portable.py
```

## 贡献

欢迎提出问题 (Issues) 或拉取请求 (Pull Requests)。

## 许可证

本项目采用 [MIT](LICENSE) 许可证。
