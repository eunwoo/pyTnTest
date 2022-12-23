import numpy as np
import matplotlib.pyplot as plt

class Sync:
    def __init__(self):
        self.n = 50
        self.m = 100
        self.y = np.linspace(-10, 10, self.n)
        self.x = np.linspace(-10, 10, self.m)
        self.phase = 0

        for i in range(self.n):
            print(self.y[i])
            yi = np.array([self.y[i]] * self.m)
            print(yi)
            d = np.sqrt(self.x ** 2 + yi ** 2)
            print(d)
            break
            z = 10 * np.cos(d + self.phase) / (d + 1)
            pts = np.vstack([self.x, yi, z]).transpose()

# s = Sync()
x = np.linspace(-10, 10, 1000)
y = np.cos(x)/(np.abs(x)+1)
print(y)
fig, ax = plt.subplots()
ax.plot(x, y)
plt.show()