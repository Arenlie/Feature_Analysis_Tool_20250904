import numpy as np
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# ===== 数据生成 =====
# 横轴频率（kHz）
freq = np.linspace(0.25, 1.75, 100)
# 纵轴细分程度（9到0方向）
granularity = np.linspace(9, 0, 100)
# 生成网格
F, G = np.meshgrid(freq, granularity)

# 模拟冲击显度分布
Z = np.zeros_like(F)
# 添加高斯峰值（模拟最大值）
peak_value = 0.44
peak_freq = 0.581  # 581Hz
peak_gran = 5.0  # 对应纵轴位置
Z += peak_value * np.exp(-((F - peak_freq) ** 2 / 0.005 + (G - peak_gran) ** 2 / 0.5))

# 添加有效值区域（559-645Hz）
mask = (F >= 0.559) & (F <= 0.645) & (G >= 3) & (G <= 7)
Z[mask] += 0.15
Z = np.clip(Z, 0, peak_value)  # 数值限幅

# ===== 可视化 =====
plt.figure(figsize=(12, 8))

# 主图绘制
im = plt.imshow(Z, cmap='magma',
                extent=[0.25, 1.75, 9, 0],  # xmin,xmax,ymax,ymin
                aspect='auto', origin='upper')
plt.colorbar(im, label='冲击显度')

# 标注最大值
plt.scatter(peak_freq, peak_gran, s=100,
            edgecolors='white', facecolors='none',
            linewidths=2, label=f'Max: {peak_value} @ {int(peak_freq * 1000)}Hz')

# 有效值区域框
plt.plot([0.559, 0.645, 0.645, 0.559, 0.559],
         [3, 3, 7, 7, 3], 'lime', linestyle='--',
         linewidth=2, label='有效值区域')

# 文字标注
text_content = """
时域冲击值: 0.0
冲击度×频次: 1.27×2.5
"""
plt.text(1.0, 2, text_content, color='white',
         bbox=dict(facecolor='black', alpha=0.8))

# 坐标设置
plt.xlabel('频率 [kHz]', fontsize=12)
plt.ylabel('频率细分程度', fontsize=12)
plt.title('20240401135110-01-001_22-34_100-2000声频冲击分布图', pad=20)
plt.legend(loc='lower right')

plt.tight_layout()
plt.savefig('output.png')
plt.show()
