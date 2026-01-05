import pandas as pd
import matplotlib.pyplot as plt
from glob import glob
import os

# Чтение всех Excel файлов
files = glob('data/*.xlsx')
all_data = []

for f in files:
    df = pd.read_excel(f)
    all_data.append((os.path.basename(f), df))  # сохраняем имя файла и DataFrame

if not all_data:
    raise ValueError("В папке 'data' нет файлов .xlsx для объединения!")

# Объединение всех данных для итогового Excel
combined = pd.concat([df for name, df in all_data], ignore_index=True)
combined.to_excel('Объединённый_отчет.xlsx', index=False)

# Итоговые суммы
summary = combined.sum(numeric_only=True)
summary_df = pd.DataFrame(summary, columns=['Итог']).T
summary_df.to_excel('Итоги_отчета.xlsx', index=False)

# Красивый график с разными линиями для каждого файла
plt.figure(figsize=(10,6))

colors = ['royalblue', 'darkorange', 'green', 'red', 'purple']  # цвета для линий

for i, (name, df) in enumerate(all_data):
    plt.plot(df['Дата'], df['Продажи'], marker='o', linewidth=2, color=colors[i % len(colors)], label=name)

plt.title('Динамика продаж по источникам', fontsize=16)
plt.xlabel('Дата', fontsize=12)
plt.ylabel('Сумма продаж', fontsize=12)
plt.xticks(rotation=45)
plt.grid(True, linestyle='--', alpha=0.7)
plt.legend(title='Файл Excel')
plt.tight_layout()
plt.savefig('График_отчета_понятный.png', dpi=300)
plt.show()
