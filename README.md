# 4G5G指标自动通报 - 安装说明

## 仓库内容

| 文件 | 说明 |
|------|------|
| full_process.py | 核心执行脚本（含指标情况计算规则） |
| 4G5G指标自动通报.skill | MobileClaw Skill导入文件 |
| README.md | 本文件 |

---

## 在其他电脑使用步骤

### 步骤1：配置目录
确保以下目录存在：
- `C:\zhibiao\4G_source` - 4G源数据目录
- `C:\zhibiao\5G_source` - 5G源数据目录
- `C:\zhibiao\4G_output` - 4G输出目录
- `C:\zhibiao\5G_output` - 5G输出目录
- `C:\zhibiao\pic_result` - 看板输出目录

### 步骤2：导入Skill
1. 将 `4G5G指标自动通报.skill` 文件复制到目标电脑
2. 在MobileClaw中导入该Skill

### 步骤3：执行
在MobileClaw中呼叫：「执行4G5G指标自动通报」

---

## 命令行执行（备用）

```bash
cd [工作目录]
uv run --with pandas --with matplotlib --with pillow --with openpyxl python full_process.py
```

---

## 指标情况计算规则

| 判定类型 | 条件 | 阈值 | 字段 |
|----------|------|------|------|
| 高负荷小区 | 上行PRB利用率≥70% 或 下行PRB利用率≥70% | ≥70% | 上行PRB利用率-%、下行PRB利用率-% |
| 低接通小区 | 无线接通率 < 95% | <95% | RRC连接建立成功率_邻区干扰-H(%) |
| 高掉线小区 | 无线掉线率 > 5% | >5% | 上行掉线率(%) |
| 低切换小区 | 切换成功率 < 90% | <90% | 切换成功率-D |
| 性能劣化小区数 | = 低接通 + 高掉线 + 低切换 | 去重 | 按「小区名称」或「基站名称」去重 |

---

## 版本信息

- 更新日期：2026-06-10
- GitHub：https://github.com/greensky113/4-5Gzhibiao