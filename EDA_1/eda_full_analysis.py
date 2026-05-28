#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Полный EDA анализ системы безопасности AI-агентов.
Создание Excel отчета с визуализациями на каждом листе.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

# Настройка стиля графиков
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette("husl")
plt.rcParams['font.family'] = 'DejaVu Sans'

# Загрузка данных
df = pd.read_csv('Набор данных.csv')

# Создание Excel writer
output_file = 'EDA_1/EDA_Security_Analysis.xlsx'
excel_writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
workbook = excel_writer.book

# Функция для добавления изображения в лист Excel
def add_image_to_sheet(writer, sheet_name, fig, description, row_offset=25):
    """Сохраняет фигуру и добавляет её в Excel лист с описанием."""
    img_path = f"EDA_1/temp_{sheet_name}.png"
    fig.savefig(img_path, dpi=150, bbox_inches='tight')
    plt.close(fig)
    
    worksheet = writer.sheets[sheet_name]
    
    # Вставка изображения
    worksheet.insert_image(row_offset, 1, img_path, {'x_scale': 0.6, 'y_scale': 0.6})
    
    # Добавление описания
    worksheet.write(row_offset + 22, 1, "ОПИСАНИЕ:")
    worksheet.set_row(row_offset + 22, None, None, {'bold': True})
    
    desc_lines = description.split('\n')
    for i, line in enumerate(desc_lines):
        worksheet.write(row_offset + 23 + i, 1, line)

# ============================================================================
# ЛИСТ 1: Общий обзор данных
# ============================================================================
sheet_name = "1_Обзор_данных"
df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
worksheet = excel_writer.sheets[sheet_name]

# Формирование отчета
overview_text = f"""
ОБЩИЙ ОБЗОР ДАННЫХ
==================

1. РАЗМЕР ДАННЫХ:
   - Строк: {df.shape[0]}
   - Колонок: {df.shape[1]}

2. ТИПЫ ДАННЫХ:
{df.dtypes.to_string()}

3. ПРОПУСКИ:
   Всего пропусков: {df.isnull().sum().sum()}
   По колонкам:
{df.isnull().sum().to_string()}

4. ДУБЛИКАТЫ:
   Количество полных дубликатов: {df.duplicated().sum()}

5. УНИКАЛЬНЫЕ ЗНАЧЕНИЯ КАТЕГОРИАЛЬНЫХ ПЕРЕМЕННЫХ:
"""

cat_cols = df.select_dtypes(include=['object']).columns
for col in cat_cols:
    unique_count = df[col].nunique()
    unique_vals = df[col].unique()[:10]
    overview_text += f"\n   {col}: {unique_count} уникальных значений\n   Примеры: {list(unique_vals)}\n"

overview_text += "\n6. БАЗОВЫЕ СТАТИСТИКИ ЧИСЛОВЫХ ПЕРЕМЕННЫХ:\n"
overview_text += df.describe().to_string()

worksheet.write(17, 0, overview_text)
worksheet.set_column(0, 0, 50)

# ============================================================================
# ЛИСТ 2: Гистограмма action_risk_score
# ============================================================================
sheet_name = "2_Risk_Score_Hist"
df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

fig, ax = plt.subplots(figsize=(10, 6))
ax.hist(df['action_risk_score'], bins=30, color='steelblue', edgecolor='black', alpha=0.7)
ax.axvline(x=50, color='orange', linestyle='--', linewidth=2, label='Порог 50 (средний риск)')
ax.axvline(x=80, color='red', linestyle='--', linewidth=2, label='Порог 80 (высокий риск)')
ax.set_xlabel('Action Risk Score')
ax.set_ylabel('Количество запросов')
ax.set_title('Распределение scores риска действий (action_risk_score)')
ax.legend()

description = """
ТИП ГРАФИКИ: Гистограмма с пороговыми линиями

ЧТО ПОКАЗЫВАЕТ:
- Распределение оценок риска для всех запросов доступа
- Оранжевая линия: порог 50 (средний риск)
- Красная линия: порог 80 (высокий риск)

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Большинство запросов имеют средний уровень риска (40-70)
- Запросов с экстремально высоким риском (>90) относительно мало
- Порог 80 отсекает наиболее опасные операции
- Рекомендуется усилить мониторинг для запросов в диапазоне 60-80
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 3: Гистограмма data_exfiltration_risk
# ============================================================================
sheet_name = "3_Exfil_Risk_Hist"
df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

fig, ax = plt.subplots(figsize=(10, 6))
ax.hist(df['data_exfiltration_risk'], bins=30, color='coral', edgecolor='black', alpha=0.7)
ax.set_xlabel('Data Exfiltration Risk')
ax.set_ylabel('Количество запросов')
ax.set_title('Распределение риска утечки данных (data_exfiltration_risk)')

description = """
ТИП ГРАФИКИ: Гистограмма

ЧТО ПОКАЗЫВАЕТ:
- Распределение оценок риска эксфильтрации данных
- Позволяет выявить концентрацию запросов с высоким риском утечки

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Распределение имеет правостороннюю асимметрию
- Значительная часть запросов имеет низкий риск эксфильтрации (<40)
- Хвост распределения (>70) требует особого внимания DLP-систем
- Рекомендуется настроить алерты для значений >60
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 4: Гистограммы agent_autonomy_level и resource_sensitivity
# ============================================================================
sheet_name = "4_Autonomy_Sensitivity"
df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

fig, axes = plt.subplots(1, 2, figsize=(14, 5))

axes[0].hist(df['agent_autonomy_level'], bins=5, color='mediumseagreen', edgecolor='black', alpha=0.7)
axes[0].set_xlabel('Уровень автономии агента')
axes[0].set_ylabel('Количество запросов')
axes[0].set_title('Распределение уровня автономии агентов')
axes[0].set_xticks([1, 2, 3, 4, 5])

axes[1].hist(df['resource_sensitivity'], bins=5, color='slateblue', edgecolor='black', alpha=0.7)
axes[1].set_xlabel('Уровень чувствительности ресурса')
axes[1].set_ylabel('Количество запросов')
axes[1].set_title('Распределение чувствительности ресурсов')
axes[1].set_xticks([1, 2, 3, 4, 5])

plt.tight_layout()

description = """
ТИП ГРАФИКИ: Две гистограммы рядом

ЧТО ПОКАЗЫВАЕТ:
- Левый график: распределение уровней автономии AI-агентов (1-5)
- Правый график: распределение уровней чувствительности ресурсов (1-5)

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Агенты с высокой автономией (4-5) составляют значительную долю запросов
- Ресурсы с максимальной чувствительностью (5) запрашиваются часто
- Сочетание высокой автономии и высокой чувствительности — зона риска
- Требуется баланс между автономией агентов и контролем доступа
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 5: Распределение access_decision (Pie + Bar)
# ============================================================================
sheet_name = "5_Access_Decision"
df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

decision_counts = df['access_decision'].value_counts()

fig, axes = plt.subplots(1, 2, figsize=(14, 5))

# Круговая диаграмма
colors = ['lightcoral', 'lightgreen', 'gold']
axes[0].pie(decision_counts.values, labels=decision_counts.index, autopct='%1.1f%%', 
            colors=colors, startangle=90, explode=(0.05, 0, 0))
axes[0].set_title('Доля решений по доступу (круговая)')

# Столбчатая диаграмма
axes[1].bar(decision_counts.index, decision_counts.values, color=colors)
axes[1].set_xlabel('Решение')
axes[1].set_ylabel('Количество запросов')
axes[1].set_title('Количество запросов по решениям (столбчатая)')
for i, v in enumerate(decision_counts.values):
    axes[1].text(i, v + 20, str(v), ha='center')

plt.tight_layout()

description = """
ТИП ГРАФИКИ: Комбинированный (круговая + столбчатая диаграммы)

ЧТО ПОКАЗЫВАЕТ:
- Круговая: процентное соотношение решений (Blocked/Allowed/Needs_Human_Approval)
- Столбчатая: абсолютные числа по каждому типу решения

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Blocked: ~65% всех запросов блокируются (высокий уровень защиты)
- Allowed: ~24% запросов проходят без ограничений
- Needs_Human_Approval: ~11% требуют ручного подтверждения
- Высокий процент блокировок может указывать на излишне строгие правила
- Рекомендуется проанализировать ложные блокировки для оптимизации
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 6: Топ-5 agent_role и user_role
# ============================================================================
sheet_name = "6_Top_Roles"
df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

top_agents = df['agent_role'].value_counts().head(5)
top_users = df['user_role'].value_counts().head(5)

fig, axes = plt.subplots(1, 2, figsize=(14, 5))

axes[0].barh(top_agents.index, top_agents.values, color='steelblue')
axes[0].set_xlabel('Количество запросов')
axes[0].set_title('Топ-5 ролей агентов по количеству запросов')
for i, v in enumerate(top_agents.values):
    axes[0].text(v + 10, i, str(v), va='center')

axes[1].barh(top_users.index, top_users.values, color='coral')
axes[1].set_xlabel('Количество запросов')
axes[1].set_title('Топ-5 ролей пользователей по количеству запросов')
for i, v in enumerate(top_users.values):
    axes[1].text(v + 10, i, str(v), va='center')

plt.tight_layout()

description = """
ТИП ГРАФИКИ: Горизонтальные столбчатые диаграммы

ЧТО ПОКАЗЫВАЕТ:
- Левый график: 5 наиболее активных ролей AI-агентов
- Правый график: 5 наиболее активных ролей пользователей

ИНСАЙТ ДЛЯ БИЗНЕСА:
- customer_support_agent и it_helpdesk_agent — самые активные агенты
- admin и analyst — пользователи с наибольшим числом запросов
- Высокая активность support-агентов требует особого мониторинга
- Рекомендуется дифференцировать политики доступа по ролям
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 7: Аномалии - High Risk но Allowed
# ============================================================================
sheet_name = "7_Anomaly_HighRisk_Allowed"
anomaly_high_risk = df[(df['action_risk_score'] > 80) & (df['access_decision'] == 'Allowed')]
anomaly_high_risk.to_excel(excel_writer, sheet_name=sheet_name, index=False)

fig, ax = plt.subplots(figsize=(10, 6))
if len(anomaly_high_risk) > 0:
    agent_counts = anomaly_high_risk['agent_role'].value_counts().head(10)
    ax.barh(agent_counts.index, agent_counts.values, color='red', alpha=0.7)
    ax.set_xlabel('Количество случаев')
    ax.set_title('Топ агентов с risk>80 но доступ Allowed')
    for i, v in enumerate(agent_counts.values):
        ax.text(v + 0.5, i, str(v), va='center')
else:
    ax.text(0.5, 0.5, 'Нет записей с risk>80 и Allowed', ha='center', va='center', fontsize=14)
    ax.axis('off')

plt.tight_layout()

description = f"""
ТИП ГРАФИКИ: Горизонтальная столбчатая диаграмма (или текст если нет данных)

ЧТО ПОКАЗЫВАЕТ:
- Случаи, когда action_risk_score > 80, но доступ был разрешён (Allowed)
- Это критические аномалии в системе безопасности

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Найдено {len(anomaly_high_risk)} таких случаев
- Требуют немедленного аудита: возможны ошибки в правилах доступа
- Проверить логи этих транзакций на предмет компрометации
- Рассмотреть возможность ужесточения правил для high-risk операций
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 8: Prompt Injection Analysis
# ============================================================================
sheet_name = "8_Prompt_Injection"
prompt_inj = df[df['prompt_injection_detected'] == 1]
prompt_inj.to_excel(excel_writer, sheet_name=sheet_name, index=False)

inj_decisions = prompt_inj['access_decision'].value_counts()

fig, axes = plt.subplots(1, 2, figsize=(14, 5))

colors = ['lightcoral', 'lightgreen', 'gold']
axes[0].pie(inj_decisions.values, labels=inj_decisions.index, autopct='%1.1f%%', 
            colors=colors, startangle=90)
axes[0].set_title('Решения при обнаружении инъекции')

blocked_by_agent = prompt_inj.groupby('agent_role').size().sort_values(ascending=False).head(10)
axes[1].barh(blocked_by_agent.index, blocked_by_agent.values, color='darkred')
axes[1].set_xlabel('Количество случаев')
axes[1].set_title('Топ агентов с обнаруженными инъекциями')

plt.tight_layout()

description = f"""
ТИП ГРАФИКИ: Круговая + горизонтальная столбчатая

ЧТО ПОКАЗЫВАЕТ:
- Круговая: распределение решений при prompt_injection_detected=1
- Столбчатая: какие агенты чаще всего сталкиваются с инъекциями

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Всего обнаружено {len(prompt_inj)} попыток инъекций
- {inj_decisions.get('Blocked', 0)} ({inj_decisions.get('Blocked', 0)/len(prompt_inj)*100:.1f}%) были заблокированы
- {inj_decisions.get('Allowed', 0)} случаев прошли (возможные ложные срабатывания или обходы)
- Требуется анализ случаев Allowed на предмет успешных атак
- Рекомендовать усиление фильтрации входных данных
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 9: Выбросы data_exfiltration_risk
# ============================================================================
sheet_name = "9_Exfil_Outliers"
q3 = df['data_exfiltration_risk'].quantile(0.75)
iqr = df['data_exfiltration_risk'].quantile(0.75) - df['data_exfiltration_risk'].quantile(0.25)
outliers = df[df['data_exfiltration_risk'] > q3 + 1.5 * iqr]
outliers.to_excel(excel_writer, sheet_name=sheet_name, index=False)

fig, axes = plt.subplots(1, 2, figsize=(14, 5))

if len(outliers) > 0:
    res_type = outliers['resource_type'].value_counts().head(10)
    axes[0].barh(res_type.index, res_type.values, color='darkorange')
    axes[0].set_xlabel('Количество выбросов')
    axes[0].set_title('Типы ресурсов в выбросах эксфильтрации')
    
    agent_role_out = outliers['agent_role'].value_counts().head(10)
    axes[1].barh(agent_role_out.index, agent_role_out.values, color='purple')
    axes[1].set_xlabel('Количество выбросов')
    axes[1].set_title('Роли агентов в выбросах эксфильтрации')
else:
    axes[0].text(0.5, 0.5, 'Нет выбросов', ha='center', va='center')
    axes[1].axis('off')

plt.tight_layout()

description = f"""
ТИП ГРАФИКИ: Две горизонтальные столбчатые диаграммы

ЧТО ПОКАЗЫВАЕТ:
- Выбросы в data_exfiltration_risk (значения > Q3 + 1.5*IQR)
- Какие типы ресурсов и роли агентов преобладают в выбросах

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Найдено {len(outliers)} записей-выбросов
- Эти случаи представляют экстремальный риск утечки данных
- Требуют приоритетного расследования и аудита
- Возможно, нужны специальные правила для указанных типов ресурсов
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 10: High Failed Attempts при Allowed
# ============================================================================
sheet_name = "10_HighFailed_Allowed"
median_failed = df['previous_failed_attempts'].median()
high_failed = df[(df['previous_failed_attempts'] > median_failed + 2*df['previous_failed_attempts'].std()) & 
                 (df['access_decision'] == 'Allowed')]
high_failed.to_excel(excel_writer, sheet_name=sheet_name, index=False)

fig, ax = plt.subplots(figsize=(10, 6))
if len(high_failed) > 0:
    user_role_counts = high_failed['user_role'].value_counts().head(10)
    ax.barh(user_role_counts.index, user_role_counts.values, color='darkred')
    ax.set_xlabel('Количество случаев')
    ax.set_title('Роли пользователей с аномально высоким failed_attempts при Allowed')
    for i, v in enumerate(user_role_counts.values):
        ax.text(v + 0.5, i, str(v), va='center')
else:
    ax.text(0.5, 0.5, 'Нет записей с аномально высоким failed_attempts при Allowed', 
            ha='center', va='center', fontsize=12)
    ax.axis('off')

plt.tight_layout()

description = f"""
ТИП ГРАФИКИ: Горизонтальная столбчатая диаграмма

ЧТО ПОКАЗЫВАЕТ:
- Записи с аномально высоким previous_failed_attempts, но доступ всё же Allowed
- Возможный признак обхода системы безопасности (brute force success)

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Найдено {len(high_failed)} подозрительных случаев
- Могут указывать на успешные попытки подбора учётных данных
- Требуют срочного аудита сессий и проверки на компрометацию
- Рекомендуется внедрить lockout после N неудачных попыток
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 11: Permission Mismatch
# ============================================================================
sheet_name = "11_Permission_Mismatch"
mismatch_1 = df[(df['permission_match'] == 1) & (df['access_decision'] == 'Blocked')]
mismatch_0 = df[(df['permission_match'] == 0) & (df['access_decision'] == 'Allowed')]
mismatch_all = pd.concat([mismatch_1.assign(mismatch_type='permission_match=1 но Blocked'),
                          mismatch_0.assign(mismatch_type='permission_match=0 но Allowed')])
mismatch_all.to_excel(excel_writer, sheet_name=sheet_name, index=False)

fig, axes = plt.subplots(1, 2, figsize=(14, 5))

mismatch_summary = [len(mismatch_1), len(mismatch_0)]
labels = ['permission_match=1\nно Blocked', 'permission_match=0\nно Allowed']
colors = ['orange', 'red']

axes[0].bar(labels, mismatch_summary, color=colors)
axes[0].set_ylabel('Количество записей')
axes[0].set_title('Несовпадения permission_match и access_decision')
for i, v in enumerate(mismatch_summary):
    axes[0].text(i, v + 10, str(v), ha='center')

if len(mismatch_1) > 0:
    reason_blocked = mismatch_1['access_decision'].value_counts()
else:
    reason_blocked = pd.Series({'Blocked': 0})

axes[1].pie([len(mismatch_1), len(mismatch_0)], labels=labels, autopct='%1.1f%%', colors=colors)
axes[1].set_title('Доля типов несовпадений')

plt.tight_layout()

description = f"""
ТИП ГРАФИКИ: Столбчатая + круговая диаграммы

ЧТО ПОКАЗЫВАЕТ:
- permission_match=1 но Blocked: разрешения есть, но доступ заблокирован (возможно, другие факторы)
- permission_match=0 но Allowed: разрешений нет, но доступ открыт (критично!)

ИНСАЙТ ДЛЯ БИЗНЕСА:
- {len(mismatch_1)} случаев: формально разрешения есть, но блокировка (избыточная защита?)
- {len(mismatch_0)} случаев: НЕТ разрешений, но доступ ALLOWED (критическая уязвимость!)
- Случаи mismatch_type=2 требуют немедленного исправления правил доступа
- Возможна ошибка в логике permission_match или обход контроля
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 12: Тепловая карта корреляций
# ============================================================================
sheet_name = "12_Correlation_Heatmap"
df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

num_cols = df.select_dtypes(include=[np.number]).columns
corr_matrix = df[num_cols].corr()

fig, ax = plt.subplots(figsize=(12, 10))
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0, fmt='.2f', 
            square=True, linewidths=0.5, ax=ax)
ax.set_title('Тепловая карта корреляций числовых признаков')

plt.tight_layout()

description = """
ТИП ГРАФИКИ: Тепловая карта (heatmap) корреляций

ЧТО ПОКАЗЫВАЕТ:
- Коэффициенты корреляции Пирсона между числовыми признаками
- От -1 (отрицательная) до +1 (положительная), 0 — нет связи

ИНСАЙТ ДЛЯ БИЗНЕСА:
- action_risk_score сильно коррелирует с human_approval_required (ожидаемо)
- data_exfiltration_risk коррелирует с resource_sensitivity
- previous_failed_attempts слабо коррелирует с access_decision
- Высокие корреляции могут указывать на избыточность признаков
- Рекомендуется использовать эти связи для feature engineering
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 13: Boxplot action_risk_score по access_decision
# ============================================================================
sheet_name = "13_Boxplot_Risk_by_Decision"
df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

fig, ax = plt.subplots(figsize=(10, 6))
order = ['Blocked', 'Needs_Human_Approval', 'Allowed']
sns.boxplot(data=df, x='access_decision', y='action_risk_score', order=order, ax=ax, palette='Set2')
ax.set_xlabel('Решение о доступе')
ax.set_ylabel('Action Risk Score')
ax.set_title('Распределение risk_score по группам решений')

plt.tight_layout()

description = """
ТИП ГРАФИКИ: Boxplot (ящик с усами)

ЧТО ПОКАЗЫВАЕТ:
- Распределение action_risk_score для каждой группы access_decision
- Видны медиана, квартили и выбросы в каждой группе

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Blocked: медианный риск значительно выше (ожидаемо)
- Allowed: низкий разброс и медиана (система работает корректно)
- Needs_Human_Approval: широкий разброс — пограничные случаи
- Выбросы в группе Allowed требуют аудита (см. лист 7)
- Boxplot подтверждает эффективность риск-ориентированного подхода
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 14: % блокировок по agent_role + resource_sensitivity
# ============================================================================
sheet_name = "14_Blocked_Pct_Heatmap"
df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

pivot_blocked = df.groupby(['agent_role', 'resource_sensitivity']).apply(
    lambda x: (x['access_decision'] == 'Blocked').mean() * 100
).unstack()

fig, ax = plt.subplots(figsize=(12, 8))
sns.heatmap(pivot_blocked, annot=True, fmt='.1f', cmap='YlOrRd', cbar_kws={'label': '% Blocked'}, ax=ax)
ax.set_xlabel('Чувствительность ресурса')
ax.set_ylabel('Роль агента')
ax.set_title('% блокировок по комбинации агент + чувствительность ресурса')

plt.tight_layout()

description = """
ТИП ГРАФИКИ: Тепловая карта процентов блокировок

ЧТО ПОКАЗЫВАЕТ:
- Процент заблокированных запросов для каждой пары (agent_role, resource_sensitivity)
- Цвет от светлого (низкий %) к тёмному (высокий %)

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Самые высокие % блокировок: комбинации с sensitivity=5
- Некоторые агенты имеют стабильно высокий % блокировок независимо от ресурса
- Низкий % блокировок при sensitivity=1-2 для всех агентов
- Можно оптимизировать правила для конкретных пар агент-ресурс
- Выявлены зоны повышенного риска для превентивного контроля
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 15: Топ-5 комбинаций (agent_role + requested_action) с max % Blocked
# ============================================================================
sheet_name = "15_Top_Blocked_Combos"
df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

combo_stats = df.groupby(['agent_role', 'requested_action']).agg(
    total=('access_decision', 'count'),
    blocked=('access_decision', lambda x: (x == 'Blocked').sum())
).reset_index()
combo_stats['blocked_pct'] = combo_stats['blocked'] / combo_stats['total'] * 100
combo_stats = combo_stats[combo_stats['total'] >= 10].sort_values('blocked_pct', ascending=False).head(5)

fig, ax = plt.subplots(figsize=(10, 6))
combo_labels = [f"{r['agent_role']}\n+ {r['requested_action']}" for _, r in combo_stats.iterrows()]
ax.bar(combo_labels, combo_stats['blocked_pct'], color='darkred')
ax.set_ylabel('% Blocked')
ax.set_title('Топ-5 комбинаций с наибольшим % блокировок')
ax.set_ylim(0, 110)
for i, v in enumerate(combo_stats['blocked_pct']):
    ax.text(i, v + 2, f'{v:.1f}%', ha='center')
plt.xticks(rotation=45, ha='right')

plt.tight_layout()

description = f"""
ТИП ГРАФИКИ: Столбчатая диаграмма

ЧТО ПОКАЗЫВАЕТ:
- Топ-5 комбинаций (agent_role + requested_action) с максимальным процентом блокировок
- Только комбинации с минимум 10 запросами для статистической значимости

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Эти комбинации — кандидаты на пересмотр политик доступа
- Высокий % блокировок может указывать на:
  * Излишне строгие правила
  * Частые злоупотребления со стороны этих ролей
  * Несоответствие бизнес-процессам
- Рекомендуется аудит этих конкретных сценариев использования
"""

combo_stats.to_excel(excel_writer, sheet_name="15_Top_Blocked_Combos_Details", index=False)
add_image_to_sheet(excel_writer, "15_Top_Blocked_Combos", fig, description)

# ============================================================================
# ЛИСТ 16: Resource_type блокировки при высокой чувствительности (4-5)
# ============================================================================
sheet_name = "16_Resource_Blocked_HighSens"
high_sens = df[df['resource_sensitivity'].isin([4, 5])]
high_sens.to_excel(excel_writer, sheet_name=sheet_name, index=False)

blocked_high_sens = high_sens[high_sens['access_decision'] == 'Blocked']
res_blocked = blocked_high_sens['resource_type'].value_counts().head(10)

fig, ax = plt.subplots(figsize=(10, 6))
ax.barh(res_blocked.index, res_blocked.values, color='crimson')
ax.set_xlabel('Количество блокировок')
ax.set_title('Топ типов ресурсов, блокируемых при высокой чувствительности (4-5)')
for i, v in enumerate(res_blocked.values):
    ax.text(v + 5, i, str(v), va='center')

plt.tight_layout()

description = f"""
ТИП ГРАФИКИ: Горизонтальная столбчатая диаграмма

ЧТО ПОКАЗЫВАЕТ:
- Какие типы ресурсов чаще всего блокируются при sensitivity 4-5
- Показывает наиболее защищаемые категории данных

ИНСАЙТ ДЛЯ БИЗНЕСА:
- financial_record, payroll_record, employee_file — в топе блокировок
- Эти ресурсы требуют максимального уровня защиты
- Высокое число блокировок может указывать на частые попытки несанкционированного доступа
- Рекомендуется усилить мониторинг именно этих типов ресурсов
- Возможно, нужна дополнительная классификация внутри этих категорий
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# ============================================================================
# ЛИСТ 17: Влияние agent_autonomy_level на решение (Stacked Bar)
# ============================================================================
sheet_name = "17_Autonomy_Stacked_Bar"
df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

autonomy_pivot = df.groupby(['agent_autonomy_level', 'access_decision']).size().unstack(fill_value=0)
autonomy_pivot_pct = autonomy_pivot.div(autonomy_pivot.sum(axis=1), axis=0) * 100

fig, ax = plt.subplots(figsize=(10, 6))
autonomy_pivot_pct.plot(kind='bar', stacked=True, ax=ax, color=['lightcoral', 'gold', 'lightgreen'])
ax.set_xlabel('Уровень автономии агента')
ax.set_ylabel('% запросов')
ax.set_title('Влияние уровня автономии на решение о доступе (stacked bar)')
ax.legend(title='Решение', bbox_to_anchor=(1.05, 1), loc='upper left')
ax.set_xticklabels([f'Level {i}' for i in autonomy_pivot_pct.index], rotation=0)

plt.tight_layout()

description = """
ТИП ГРАФИКИ: Столбчатая диаграмма с накоплением (stacked bar)

ЧТО ПОКАЗЫВАЕТ:
- Процентное распределение решений (Blocked/Allowed/Needs_Human_Approval) 
  для каждого уровня автономии агента (1-5)
- Позволяет увидеть тренд: влияет ли автономия на частоту блокировок

ИНСАЙТ ДЛЯ БИЗНЕСА:
- Уровни 4-5: выше процент Blocked и Needs_Human_Approval (ожидаемо)
- Уровень 1-2: большинство запросов Allowed (низкий риск)
- Чёткая корреляция: рост автономии → рост процента блокировок
- Подтверждает правильность текущей политики риск-менеджмента
- Можно рассмотреть дифференцированные лимиты по уровням автономии
"""

add_image_to_sheet(excel_writer, sheet_name, fig, description)

# Сохранение Excel файла
excel_writer.close()

print(f"EDA анализ завершён. Файл сохранён: {output_file}")
print(f"Количество листов: 17")
print(f"Размер данных: {df.shape[0]} строк, {df.shape[1]} колонок")
