import matplotlib.pyplot as plt
import numpy as np

# 產生一些示例數據
x = np.linspace(0, 2 * np.pi, 100)
y = np.sin(x) + np.random.normal(0, 0.1, size=len(x))

# 使用 numpy 中的 smooth 函數實現平滑
smoothed_y = np.convolve(y, np.ones(10)/10, mode='valid')

# 繪製原始折線圖
plt.plot(x, y, label='original')

# 繪製平滑折線圖
plt.plot(x[:len(smoothed_y)], smoothed_y, label='smoothed')

plt.xlabel('X')
plt.ylabel('Y')
plt.title('title')
plt.legend()
plt.show()