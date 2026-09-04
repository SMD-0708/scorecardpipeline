# -*- coding: utf-8 -*-
"""
新功能示例：SHAP 解释性、模型监控、概率校准

本脚本展示 scorecardpipeline 新增的三大能力：
1. 使用 ScorecardExplainer 对 WOE-LR 评分卡进行 SHAP 解释
2. 使用 ModelMonitor 监控评分分布与特征漂移
3. 使用 ProbabilityCalibrator 对模型输出概率进行校准
"""

from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline

from scorecardpipeline import (
    FeatureSelection,
    Combiner,
    WOETransformer,
    ITLubberLogisticRegression,
    germancredit,
)
from scorecardpipeline.explainability import ScorecardExplainer
from scorecardpipeline.monitoring import ModelMonitor
from scorecardpipeline.calibration import ProbabilityCalibrator


def main():
    # 1. 准备数据
    data = germancredit()
    data["creditability"] = data["creditability"].map({"good": 0, "bad": 1})
    train, test = train_test_split(
        data, test_size=0.3, random_state=42, stratify=data["creditability"]
    )

    # 2. 构建标准评分卡 Pipeline
    pipeline = Pipeline([
        ("select", FeatureSelection(target="creditability", engine="scorecardpy")),
        ("combiner", Combiner(target="creditability", method="chi", max_n_bins=4)),
        ("woe", WOETransformer(target="creditability")),
    ])
    woe_train = pipeline.fit_transform(train)
    woe_test = pipeline.transform(test)

    model = ITLubberLogisticRegression(target="creditability")
    model.fit(woe_train)

    # 3. SHAP 解释性
    explainer = ScorecardExplainer(
        model=model,
        combiner=pipeline.named_steps["combiner"],
        woe_transformer=pipeline.named_steps["woe"],
    )
    print("单样本解释：")
    print(explainer.explain_sample(test, index=0).head())
    print("\n特征重要度汇总：")
    print(explainer.summary(test).head())

    # 4. 模型监控
    score_train = model.predict_proba(woe_train)[:, 1]
    score_test = model.predict_proba(woe_test)[:, 1]
    y_test = test["creditability"].values

    monitor = ModelMonitor(score_bins=10)
    monitor.fit_reference(
        train[pipeline.named_steps["combiner"].combiner.rules.keys()],
        score_train,
        y_true=train["creditability"].values,
    )
    print("\n评分 PSI:", monitor.score_psi(score_test))
    print("\n特征 PSI:")
    print(monitor.feature_psi(test[pipeline.named_steps["combiner"].combiner.rules.keys()]))
    print("\n性能衰减:")
    print(monitor.performance_decay(score_test, y_test))

    # 5. 概率校准
    calibrator = ProbabilityCalibrator(method="platt")
    calibrated = calibrator.fit_transform(score_train, train["creditability"].values)
    print("\n校准后概率范围:", calibrated.min(), "~", calibrated.max())


if __name__ == "__main__":
    main()
