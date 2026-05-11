# 钢铁行业减碳基准路径模型 V11.1

## 项目概述
基于 V11 版本的钢铁行业碳减排模型优化，实现三情景碳强度与 Dashboard 目标的精确对齐。

## 核心修改内容

### 1. 产量差异化（Step 3）
- 积极情景（NZE）：低产量路径，2060年61600万吨
- 适度情景（2°C CGE）：中产量路径，2060年66600万吨
- 保守情景（NDC）：高产量路径，2060年73100万吨

### 2. 废钢资源量参考（Step 5）
- 更新为中国冶金协会三情景预测数据
- 三情景差异化废钢供给曲线

### 3. 流程结构参数（Step 4）
- 短流程占比：积极14%/适度12%/保守10.5%（2030年）
- 氢冶金占比：积极2%/适度1.8%/保守1.5%（2030年）
- 各情景独立平滑递增曲线

### 4. 绿电占比（Step 6）
- 三情景差异化：积极>适度>保守
- 平滑递增，2060年达到88%/82%/72%

### 5. 低碳工艺技术（Step 7）
- 三情景差异化技术普及率
- 2040年前完全平滑（线性递增，无跳跃）
- 比例：积极1.0/适度0.85/保守0.65

### 6. CCUS动态配平（Step 8）
- Row 59改为动态公式：`=MAX(0,MIN(0.98,IF(curve_tech>0,1-Dashboard目标CI*产量/curve_tech,0)))`
- 引用Dashboard目标CI，自动响应参数变化
- 适度情景2060年CCUS≈2.79亿吨

## 最高原则验证结果

✅ **三情景全部37年（2024-2060）碳强度偏差 = 0%**
✅ 适度情景2030年降幅7.79%（目标7.8%，误差<0.01%）
✅ 适度情景2035年降幅17.59%（目标17.6%，误差<0.01%）

## 文件结构

### 输入文件
- `副本钢铁行业减碳基准路径模型_V11.xlsx` - 原始数据源
- `prd2.md` - 优化需求文档
- `钢铁行业减碳模型_知识分享文档.md` - 模型知识库
- `钢铁行业模型_修正指令.md` - Agent修正指令

### 分析脚本（按功能分类）
- `analyze_model.py` / `analyze_model_v2.py` - 初始结构分析
- `deep_analysis.py` / `deep_analysis_v3.py` - 深度参数分析
- `compare_analysis.py` - 跨情景对比分析
- `moderate_analysis.py` - 适度情景专项分析
- `check_v11_alignment.py` - V11对齐检查
- `reverse_engineer.py` - V11公式逆向工程
- `debug_scrap.py` - 废钢计算调试
- `check_dashboard_pos.py` - Dashboard位置确认
- `design_params.py` - 参数设计（逆推curve_tech）
- `analyze_strategy.py` - 核心数学策略分析

### 参数验证脚本
- `final_params_v2.py` - 最终参数设计与验证（核心脚本）
- `calc_params.py` - CCUS配平计算
- `check_remaining.py` - 剩余工作检查

### Excel生成脚本
- `build_model.py` / `build_model_v2.py` - 早期构建尝试
- `build_final.py` - 构建脚本
- `optimize_steel_model.py` - 优化尝试
- `write_excel_v11.py` - Excel写入尝试
- `write_v11_final.py` - 最终Excel参数写入（**核心脚本**）
- `add_ccus_formula.py` - CCUS动态公式替换
- `full_validate.py` - 全量验证脚本
- `verify_excel.py` - Excel重算验证

### 输出文件
- `~/Desktop/钢铁行业模型V11.1.xlsx` - 最终优化结果（已保存到桌面）

## 技术要点

### Step 8 瀑布流逻辑
```
curve_prod = P × 1.95 （产量×基准CI）
  → curve_raw  = curve_prod - 原料优化降碳（可负）
  → curve_eaf  = curve_raw  - 短流程提升降碳（可负）
  → curve_h2   = curve_eaf  - 氢冶金降碳（≥0）
  → curve_green = curve_h2  - 绿电降碳
  → curve_tech = curve_green - 低碳工艺降碳
  → curve_final = curve_tech × (1 - CCUS比例)
  → CI = curve_final / P
```

### CCUS配平核心约束
- CCUS只能减少排放（比例≥0）
- 因此：curve_tech ≥ 目标CI × P 必须成立
- 配平公式：CCUS比例 = 1 - 目标CI × P / curve_tech
- 限幅：0 ≤ CCUS比例 ≤ 0.98

### 关键数学发现
- V11原始模型三情景使用相同流程参数+相同产量，导致积极情景2030年curve_tech_CI远低于目标CI，无法配平
- 解决方案：让技术参数足够保守（短流程/H2/绿电占比平缓增长），使curve_tech始终高于目标CI×P
- CCUS负责配平差额，S型曲线确保物理可行性