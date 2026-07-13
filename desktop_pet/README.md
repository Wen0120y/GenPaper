# Codex 桌面宠物

基于 `spritesheet.webp` 和 `validation.json` 的 Windows 桌面宠物程序。

## 启动

```powershell
python pet.py
```

如果缺少 Pillow：

```powershell
python -m pip install -r requirements.txt
```

## 使用

- 左键单击角色：思考。
- 左键双击角色：跳一下，并打开 <https://actaresearch.streamlit.app/>。
- 左键按住并移动：拖动宠物。
- 右键短按：打开动作菜单。
- 右键按住并移动：拖动宠物。
- 拖动时只使用奔跑动作：`running-left` / `running-right`。
- `休息` 展示喝咖啡动作。
- `思考` 展示待机角色右上角灯泡亮起。
- `论文` 展示写代码动作。
- `实验` 展示实验动作。
- `论文`、`实验`、`休息` 可以选择 5 分钟、30 分钟、1 小时或 2 小时。

## 动作映射

- 待机：`idle`
- 单击：`thinking`
- 双击：`jumping` + 打开网页
- 拖动：`running-left` / `running-right`
- 休息：`waiting`
- 思考：`thinking`
- 论文 / 写代码：`running`
- 实验：`review`
