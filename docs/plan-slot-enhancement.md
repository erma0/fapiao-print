# 单票调整优化方案

> 版本: v2.0.2 目标 | 日期: 2026-06-01 | 审校: 2026-06-01

## 核心问题

**能放大能拖，但拖不到位。**

现有功能已支持自适应大小 + 手动放大 + 拖动偏移，理论上用户可以把发票拖到 slot 中间。但实际操作中存在三个卡点：

1. **放大不够**：slotScale 上限 2.0，地铁行程单等窄长发票 2x 仍显小
2. **拖拽范围太小**：硬编码 `maxOffY = slot.h × 0.5`，不看发票实际放大后的尺寸。放大越多超出 slot 越多，但可拖距离不变 → 拉不到位
3. **编辑时看不见溢出**：`overflow: hidden` 裁掉了超出 slot 的内容，用户不知道该往哪拖

举例：宽发票 contain-fit 后放大 2x，显示高度 700px 超出 slot（400px）300px，但拖拽范围只有 200mm → 拖不到中间。

---

## 设计原则

**发票边界 ≤ 插槽边界，不溢出。**

- 插槽（slot）= 纸张减去边距后的可印区域，是发票内容的硬边界
- PDF/打印输出中，发票内容始终被 clip 在 slot 内，无论 1×1 还是多版布局
- 前端预览中，**编辑态**（选中/拖拽）临时 `overflow:visible` 方便用户判断调整方向；**非编辑态** `overflow:hidden` 与输出一致
- 超 A5 发票的适配策略：通过 **放大 + 拖动** 将重要内容调整到 slot 可见区，边缘内容接受裁切

### 为什么不允许溢出

| 考量 | 允许溢出 | 不溢出（当前方案） |
|------|---------|------------------|
| 预览与输出一致性 | ❌ 需额外同步前后端 clip 逻辑 | ✅ 非编辑态始终一致 |
| 打印安全性 | ⚠️ 内容可能超出纸张可印区 | ✅ slot = 可印区，不会越界 |
| 多版布局 | ⚠️ 需区分 1×1/多版两套逻辑 | ✅ 统一一套逻辑 |
| 超 A5 发票 | ✅ 可见完整内容 | ⚠️ 边缘被裁切，需用户调整 |
| 实现复杂度 | ⚠️ Rust 端需条件 clip | ✅ 无 Rust 改动 |

"不溢出"牺牲了超 A5 发票的边缘完整性，但换来了一致性、安全性和简洁性。本次优化让拖拽范围跟得上放大倍数，用户可以轻松将重要内容拖到可见区。

---

## 改动清单

共 **3 项改动**，跨 **4 个文件**，改动约 25 行。

| # | 改动 | 文件 | 解决的卡点 |
|---|------|------|-----------|
| A | slotScale 上限 2.0→3.0 | `index.html` `app.js` `layout.js` | 放大不够 |
| B | 拖拽约束动态化 + 编辑态溢出可见 | `layout.js` `styles.css` | 拖不到位 + 看不见溢出 |
| C | 单击选中 + 滚轮调节单票缩放 | `app.js` | 操作不便 |

---

## 改动 A：slotScale 上限 2.0 → 3.0

### 原因

地铁行程单（窄长型，宽高比约 0.36）经 contain-fit 后被宽度约束，显示尺寸偏小。2x 放大仍然不够，存在大量 slot 空白。放宽到 3x。

### 影响位置（6 处）

| 文件 | 行号 | 位置 | 改前 | 改后 |
|------|------|------|------|------|
| `layout.js` | 399 | 拖拽缩放约束 | `Math.min(2.0)` | `Math.min(3.0)` |
| `index.html` | 140 | 单票调整面板 range | `max="200"` | `max="300"` |
| `index.html` | 140 | 单票调整面板 number | `max="200"` | `max="300"` |
| `app.js` | 1738 | `onAdjScaleChange()` | `Math.min(2.0)` | `Math.min(3.0)` |
| `app.js` | 1635 | 弹窗 HTML 模板 | `max="200"` | `max="300"` |
| `app.js` | 1663 | `confirmInvModal()` | `Math.min(2.0)` | `Math.min(3.0)` |

### 风险

- **无**。只是放宽上限，最小值和默认值不变。

---

## 改动 B：拖拽约束动态化 + 编辑态溢出可见

### 原因

**拖拽范围与放大倍数脱节**：当前硬编码 `maxOffY = slot.h × 0.5`，不看发票实际放大后的尺寸。放大越多超出 slot 越多，但可拖距离不变 → 拉不到位。

**编辑时看不见溢出内容**：`.invoice-slot { overflow: hidden }` 裁掉了超出 slot 的内容，用户无法判断调整方向和幅度。需要在编辑态临时放开，非编辑态恢复 clip 保持与输出一致。

### B1：拖拽约束动态化

**文件**：`src/layout.js` → `onSlotMouseMove()` (约第 380-384 行)

**改前**：
```javascript
var maxOffX = (slot.w / MM2PX) * 0.5;
var maxOffY = (slot.h / MM2PX) * 0.5;
newOffX = Math.max(-maxOffX, Math.min(maxOffX, newOffX));
newOffY = Math.max(-maxOffY, Math.min(maxOffY, newOffY));
```

**改后**：
```javascript
var f = _slotDrag.fileObj;
var imgW = f.ow || 1;
var imgH = f.oh || 1;
var s = _slotDrag.cachedSettings;
var displayW, displayH;
if (s.fitMode === 'fill') {
  displayW = slot.w * (f.slotScale || 1);
  displayH = slot.h * (f.slotScale || 1);
} else if (s.fitMode === 'original') {
  var rDpi = f.renderDpi || 150;
  displayW = (imgW / (rDpi / 25.4)) * MM2PX * (f.slotScale || 1);
  displayH = (imgH / (rDpi / 25.4)) * MM2PX * (f.slotScale || 1);
} else {
  // contain / custom
  var fitScale = Math.min(slot.w / imgW, slot.h / imgH);
  var perScale = f.slotScale || 1;
  displayW = imgW * fitScale * perScale;
  displayH = imgH * fitScale * perScale;
  if (s.fitMode === 'custom' && s.customScale !== 1) {
    displayW *= s.customScale;
    displayH *= s.customScale;
  }
}
var extraX = Math.max(0, (displayW - slot.w) / 2 / MM2PX);
var extraY = Math.max(0, (displayH - slot.h) / 2 / MM2PX);
var minRangeX = (slot.w / MM2PX) * 0.5;
var minRangeY = (slot.h / MM2PX) * 0.5;
var maxOffX = Math.max(minRangeX, extraX);
var maxOffY = Math.max(minRangeY, extraY);
newOffX = Math.max(-maxOffX, Math.min(maxOffX, newOffX));
newOffY = Math.max(-maxOffY, Math.min(maxOffY, newOffY));
```

**逻辑**：
- 发票放大后超出 slot 越多 → 可拖范围越大（`extraX/Y` 随溢出量增长）
- 发票没超出 slot → 至少保留原半个 slot 的可拖范围（`minRangeX/Y` 兜底）
- 兼容所有 fitMode（contain / fill / original / custom），与 `renderPage` 中 contained 尺寸计算逻辑一致

**效果对比**：

| 场景 | 改前 maxOffY | 改后 maxOffY |
|------|-------------|-------------|
| contain 1x，发票小于 slot | slot.h × 0.5 | slot.h × 0.5（不变） |
| contain 2x，发票超出 slot 100px | slot.h × 0.5 | slot.h × 0.5 + 50/MM2PX（更大） |
| contain 3x，发票超出 slot 300px | slot.h × 0.5 | slot.h × 0.5 + 150/MM2PX（大得多） |

### B2：编辑态溢出可见

**文件**：`src/styles.css` → 在 `.invoice-slot.selected` 和 `.invoice-slot.dragging` 规则后追加

```css
.invoice-slot.selected { overflow: visible !important; }
.invoice-slot.dragging { overflow: visible !important; }
```

**效果**：选中或拖拽中的发票，内容可以超出 slot 边界显示，方便用户判断调整方向和幅度。

**前后端一致性**：

| 状态 | 前端预览 | 后端 PDF | 一致？ |
|------|---------|---------|--------|
| 选中/拖拽（编辑态） | `overflow: visible` | clip 到 slot | ⚠️ 临时不一致（编辑辅助） |
| 未选中（非编辑态） | `overflow: hidden` | clip 到 slot | ✅ 一致 |

编辑态的临时不一致是**有意为之**：用户需要看到溢出内容才能判断调整方向和幅度。取消选中后前端恢复 clip，与输出一致。最终 PDF/打印始终 clip，内容不会超出 slot。

### 风险

- 多版布局（2×2 等）中选中槽溢出可能短暂遮挡相邻槽。但用户**主动选中该槽**就是想调整它，遮挡是预期的。取消选中后 `.selected` 移除，恢复 `overflow: hidden`。
- 不影响实际打印（Rust 端始终 clip）。

---

## 改动 C：单击选中 + 滚轮调节单票缩放

### 原因

当前单票缩放只存在于侧边栏面板和角手柄拖拽，用户在预览区看不到明显的缩放入口。滚轮是最直觉的缩放操作，应直接支持。

### 改法

**文件**：`src/app.js` → 现有 Ctrl+wheel handler 前方插入判断（约第 2461 行）

**新增代码**：
```javascript
if (!e.ctrlKey && S.selectedSlot >= 0) {
  var slotEl = e.target.closest('.invoice-slot');
  if (slotEl && parseInt(slotEl.dataset.slotIdx) === S.selectedSlot) {
    e.preventDefault();
    var f = getSelectedFileObj();
    if (f) {
      var step = 5;
      var curPct = Math.round((f.slotScale || 1) * 100);
      var newPct = e.deltaY > 0 ? curPct - step : curPct + step;
      f.slotScale = Math.max(0.2, Math.min(3.0, newPct / 100));
      updatePreview();
      updateAdjPanel();
      return;
    }
  }
}
```

**交互逻辑**：

| 操作 | 行为 |
|------|------|
| 无 Ctrl + 未选中槽 | 正常滚动页面 |
| 无 Ctrl + 已选中槽 + 鼠标在槽上 | 调该票 slotScale（5%/click） |
| 无 Ctrl + 已选中槽 + 鼠标不在槽上 | 正常滚动页面 |
| Ctrl + 滚轮 | 缩放整个视图（不变） |

### 风险

- **无**。只在有选中槽且鼠标在选中槽上时才拦截滚轮，不影响任何现有操作。

---

## 实施顺序

改动之间无依赖，建议顺序：

```
1. A（slotScale 上限） → 最小改动，先做
2. B（拖拽约束 + 编辑态溢出） → 与 A 配合生效
3. C（滚轮缩放） → 独立新功能
```

---

## 回滚方案

- 所有改动均为 `git stash` 可回退的纯增量
- 全部改动仅涉及 JS/HTML/CSS，**不触发 Rust 重编译**
