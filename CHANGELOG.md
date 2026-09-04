# Changelog

## [Unreleased]

### Added

- **新模块 `scorecardpipeline.explainability`**：
  - 新增 `ScorecardExplainer` 类，支持对 WOE-LR 评分卡模型进行 SHAP 解释。
  - 支持批量解释、单样本解释、特征重要度汇总。
- **新模块 `scorecardpipeline.monitoring`**：
  - 新增 `ModelMonitor` 类，支持评分分布 PSI、特征 PSI、模型性能（KS/AUC）衰减监控。
  - 提供 `monitor_report` 方法生成完整监控报告。
- **新模块 `scorecardpipeline.calibration`**：
  - 新增 `ProbabilityCalibrator` 类，支持 Platt Scaling 和 Isotonic Regression 概率校准。
  - 兼容单概率、概率矩阵等多种输入形式。
- **测试与 CI/CD**：
  - 新增 `tests/` 目录及完整回归测试（`pytest`），覆盖 processing、scorecard、model、rule、explainability、monitoring、calibration 等核心模块。
  - 新增 GitHub Actions 工作流：多平台/多版本测试（`.github/workflows/tests.yml`）和代码风格检查（`.github/workflows/lint.yml`）。
  - 新增 `.pre-commit-config.yaml`，集成 `ruff` 和 `black`。
- **依赖治理**：
  - 将 `pyproject.toml` 中的依赖拆分为 core / graph / pmml / eda / explain / dev / all 可选依赖组，降低最小化安装门槛。

### Fixed

- 修复 `Combiner.load()` 加载离线规则后未设置 `fitted_` 的问题。
- 修复 `scorecard.py` 在 `scikit-learn >= 1.6` 环境下 `check_array` 参数不兼容的问题。

### Changed

- 更新 `README.md`，补充按能力组安装命令示例。
- 更新 `CLAUDE.md`，增加新模块架构说明。
