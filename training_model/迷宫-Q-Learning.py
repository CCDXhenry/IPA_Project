import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Qt5Agg')  # 必须在导入pyplot之前设置
import matplotlib.pyplot as plt
from itertools import product
# 定义环境
maze = [
    [0, 1, 0, 0, 0],
    [0, 1, 0, 1, 0],
    [0, 0, 0, 0, 0],
    [0, 1, 1, 1, 0],
    [0, 0, 0, 0, 0]
]

actions = [(0, -1), (0, 1), (-1, 0), (1, 0)]  # 左、右、上、下
goal = (4, 4)
start = (0, 0)


def step(state, action):
    x, y = state
    dx, dy = action
    new_state = (x + dx, y + dy)

    if not (0 <= new_state[0] < len(maze)) or not (0 <= new_state[1] < len(maze[0])):
        return state, -1

    if maze[new_state[0]][new_state[1]] == 1:
        return state, -1

    reward = -0.01
    if new_state == goal:
        reward = 1
    return new_state, reward


# 初始化Q表
states = [str(state) for state in product(range(len(maze)), repeat=2)]
Q = pd.DataFrame(0.0, index=states, columns=range(len(actions)))

# Q-Learning参数
alpha = 0.1  # 学习率
gamma = 0.9  # 折扣因子
epsilon = 0.9  # 探索率

print(Q)
# 用于存储每一步的信息
states_history = []
rewards_history = []

# 训练
for episode in range(1000):
    state = start
    states_history.append(state)
    rewards_history.append(0)

    while True:
        if np.random.uniform() < epsilon:
            action_idx = np.random.choice(range(len(actions)))
        else:
            action_idx = Q.loc[str(state)].idxmax()

        next_state, reward = step(state, actions[action_idx])
        Q.loc[str(state), action_idx] += alpha * (
                    reward + gamma * Q.loc[str(next_state)].max() - Q.loc[str(state), action_idx])

        state = next_state
        states_history[-1] = state
        rewards_history[-1] += reward

        if state == goal:
            break
    print(episode)
# 可视化
plt.figure(figsize=(10, 5))

# 绘制每一步的奖励总和
plt.subplot(1, 2, 1)
plt.plot(rewards_history)
plt.title('Total Reward per Episode')
plt.xlabel('Episode')
plt.ylabel('Total Reward')

# 绘制智能体的路径
plt.subplot(1, 2, 2)
maze_plot = np.array(maze)
maze_plot[start] = 2
maze_plot[goal] = 3
for state in states_history[:10]:  # 显示前10次尝试的路径
    maze_plot[state] = 4
plt.imshow(maze_plot, cmap='gray', interpolation='nearest')
plt.colorbar()
plt.title('Agent Path')
plt.xticks([])
plt.yticks([])

plt.tight_layout()
plt.show()