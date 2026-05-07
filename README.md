# SKILL for CQUer

重庆大学微电子与通信工程学院实验报告自动化生成技能集，基于 [OpenHanako](https://github.com/liliMozi/openhanako) 平台。

## 技能列表

### cqu-random-signal-report
随机信号分析实验报告生成器（MATLAB 仿真类）。

- 读取实验要求（.docx / .pdf）
- 自动编写并运行 MATLAB 脚本
- 用学院电子化模板生成 Word 实验报告
- 支持 OMML 公式、自动目录、标准封面

### cqu-lab-report
通用实验报告生成器，适用于学院所有实验课程。

- 支持多种实验类型：MATLAB 仿真、Quartus/ModelSim、实验室实物
- 支持新建报告和增量续写两种模式
- 严格对齐学院六节标准格式（1.1实验目的 ~ 1.6实验结果及分析）
- 奇偶页眉（奇数页=实验名 / 偶数页=微电子与通信工程学院实验报告）
- 分节分页、页脚下划线、自动目录

## 使用方式

将 `SKILL.md` 文件安装到 OpenHanako 平台即可。具体安装方法参考 [OpenHanako 文档](https://github.com/liliMozi/openhanako)。

## 更新日志

### 2026-05-07
- `cqu-lab-report`: 完成格式验证，对齐标准实验报告模板
  - 标题数字与文字间无空格（`1.1实验目的`）
  - 子任务用 `1）` 编号，不进入目录
  - 图注格式 `图X.N`
  - 代码以截图形式嵌入，不设附录
  - 封面提取后清理多余分节符
  - 奇偶页眉独立 header Part

