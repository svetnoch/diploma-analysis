#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EDA анализ данных системы безопасности AI-агентов и access control
Senior Data Analyst Analysis
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
warnings.filterwarnings('ignore')

# Настройка стиля графиков
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette("husl")

# Загрузка данных
print("Загрузка данных...")
df = pd.read_csv('/workspace/Набор данных.csv')

print(f"Размер данных: {df.shape[0]} строк, {df.shape[1]} колонок")
print("\n" + "="*80 + "\n")

# ============================================
# 1. ОБЩИЙ ОБЗОР ДАННЫХ
# ============================================
print("1. ОБЩИЙ ОБЗОР ДАННЫХ")
print("="*80)

# Размер данных
print(f"\n1.1 Размер данных: {df.shape[0]} строк, {df.shape[1]} колонок")

# Типы данных
print("\n1.2 Типы данных:")
print(df.dtypes)

# Пропуски
print("\n1.3 Пропуски в данных:")
missing = df.isnull().sum()
missing_pct = (df.isnull().sum() / len(df) * 100).round(2)
missing_df = pd.DataFrame({'Пропуски': missing, 'Процент': missing_pct})
print(missing_df[missing_df['Пропуски'] > 0] if len(missing_df[missing_df['Пропуски'] > 0]) > 0 else "Пропусков нет")

# Дубликаты
duplicates = df.duplicated().sum()
print(f"\n1.4 Дубликаты: {duplicates} строк ({round(duplicates/len(df)*100, 2)}%)")

# Уникальные значения категориальных переменных
print("\n1.5 Уникальные значения категориальных переменных:")
categorical_cols = df.select_dtypes(include=['object']).columns
for col in categorical_cols:
    unique_vals = df[col].unique()
    print(f"\n{col}: {len(unique_vals)} уникальных значений")
    print(f"  Значения: {list(unique_vals[:10])}{'...' if len(unique_vals) > 10 else ''}")

# Базовые статистики числовых переменных
print("\n1.6 Базовые статистики числовых переменных:")
numeric_cols = df.select_dtypes(include=[np.number]).columns
print(df[numeric_cols].describe())

# ============================================
# СОЗДАНИЕ Excel файла с результатами
# ============================================
print("\n" + "="*80)
print("СОЗДАНИЕ Excel файла с результатами анализа")
print("="*80)

# Создаем Excel writer
excel_file = '/workspace/EDA_Security_Analysis.xlsx'

with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    
    # Лист 1: Общий обзор
    overview_data = {
        'Метрика': ['Количество строк', 'Количество колонок', 'Дубликаты', 'Дубликаты %'],
        'Значение': [df.shape[0], df.shape[1], duplicates, round(duplicates/len(df)*100, 2)]
    }
    overview_df = pd.DataFrame(overview_data)
    overview_df.to_excel(writer, sheet_name='1_Общий_обзор', index=False)
    
    # Добавляем информацию о типах данных
    types_data = pd.DataFrame({
        'Колонка': df.columns,
        'Тип': df.dtypes.values,
        'Пропуски': missing.values,
        'Пропуски %': missing_pct.values
    })
    types_data.to_excel(writer, sheet_name='1_Общий_обзор', startrow=len(overview_df)+2, index=False)
    
    # Лист 2: Уникальные значения категориальных переменных
    categorical_summary = []
    for col in categorical_cols:
        value_counts = df[col].value_counts()
        for val, count in value_counts.items():
            categorical_summary.append({
                'Переменная': col,
                'Значение': val,
                'Количество': count,
                'Процент': round(count/len(df)*100, 2)
            })
    categorical_df = pd.DataFrame(categorical_summary)
    categorical_df.to_excel(writer, sheet_name='2_Категориальные', index=False)
    
    # Лист 3: Статистики числовых переменных
    numeric_stats = df[numeric_cols].describe().transpose()
    numeric_stats.to_excel(writer, sheet_name='3_Числовые_статистики')
    
    # Лист 4: Аномалии - action_risk_score > 80 но Allowed
    high_risk_allowed = df[(df['action_risk_score'] > 80) & (df['access_decision'] == 'Allowed')]
    high_risk_allowed.to_excel(writer, sheet_name='4_Аномалии_Высокий_риск_Allowed', index=False)
    
    # Лист 5: Аномалии - prompt_injection_detected = 1
    injection_cases = df[df['prompt_injection_detected'] == 1]
    injection_summary = injection_cases.groupby('access_decision').size().reset_index(name='count')
    injection_summary.to_excel(writer, sheet_name='5_Prompt_Injection', index=False)
    
    # Лист 6: Выбросы data_exfiltration_risk
    exfil_outliers = df[df['data_exfiltration_risk'] > df['data_exfiltration_risk'].quantile(0.95)]
    exfil_outliers.to_excel(writer, sheet_name='6_Exfiltration_Outliers', index=False)
    
    # Лист 7: Аномально высокий previous_failed_attempts при успешном доступе
    high_attempts_success = df[(df['previous_failed_attempts'] > df['previous_failed_attempts'].quantile(0.9)) & 
                                (df['access_decision'] == 'Allowed')]
    high_attempts_success.to_excel(writer, sheet_name='7_High_Attempts_Allowed', index=False)
    
    # Лист 8: Несовпадения permission_match vs access_decision
    mismatch_1 = df[(df['permission_match'] == 1) & (df['access_decision'] == 'Blocked')]
    mismatch_2 = df[(df['permission_match'] == 0) & (df['access_decision'] == 'Allowed')]
    mismatch_df = pd.concat([mismatch_1.assign(issue_type='permission_match=1 но Blocked'),
                             mismatch_2.assign(issue_type='permission_match=0 но Allowed')])
    mismatch_df.to_excel(writer, sheet_name='8_Permission_Mismatches', index=False)
    
    # Лист 9: Топ комбинаций agent_role + requested_action по % Blocked
    role_action_grouped = df.groupby(['agent_role', 'requested_action']).agg({
        'access_decision': lambda x: (x == 'Blocked').sum() / len(x) * 100,
        'agent_role': 'count'
    }).rename(columns={'access_decision': '%_Blocked', 'agent_role': 'total_count'})
    role_action_grouped = role_action_grouped.sort_values('%_Blocked', ascending=False).head(5)
    role_action_grouped.to_excel(writer, sheet_name='9_Top_Blocked_Combinations')
    
    # Лист 10: Сводка по блокировкам
    blocked_summary = df.groupby(['agent_role', 'resource_sensitivity']).agg({
        'access_decision': lambda x: (x == 'Blocked').sum() / len(x) * 100
    }).pivot_table(index='agent_role', columns='resource_sensitivity', values='access_decision')
    blocked_summary.to_excel(writer, sheet_name='10_Blocked_by_Sensitivity')

print(f"\nExcel файл создан: {excel_file}")

# ============================================
# 2. ВИЗУАЛИЗАЦИИ
# ============================================
print("\n" + "="*80)
print("СОЗДАНИЕ ВИЗУАЛИЗАЦИЙ")
print("="*80)

# Настройка для сохранения графиков
fig_dir = '/workspace/visualizations'
import os
os.makedirs(fig_dir, exist_ok=True)

# --- График 1: Гистограмма action_risk_score ---
fig, ax = plt.subplots(figsize=(12, 6))
ax.hist(df['action_risk_score'], bins=30, edgecolor='black', alpha=0.7, color='steelblue')
ax.axvline(x=50, color='orange', linestyle='--', linewidth=2, label='Порог 50')
ax.axvline(x=80, color='red', linestyle='--', linewidth=2, label='Порог 80')
ax.set_xlabel('Action Risk Score', fontsize=12)
ax.set_ylabel('Количество', fontsize=12)
ax.set_title('Распределение Action Risk Score с порогами 50 и 80', fontsize=14, fontweight='bold')
ax.legend()
plt.tight_layout()
plt.savefig(f'{fig_dir}/01_action_risk_score_hist.png', dpi=150)
plt.close()
print("Создан график: 01_action_risk_score_hist.png")

# --- График 2: Гистограмма data_exfiltration_risk ---
fig, ax = plt.subplots(figsize=(12, 6))
ax.hist(df['data_exfiltration_risk'], bins=30, edgecolor='black', alpha=0.7, color='coral')
ax.set_xlabel('Data Exfiltration Risk', fontsize=12)
ax.set_ylabel('Количество', fontsize=12)
ax.set_title('Распределение Data Exfiltration Risk', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{fig_dir}/02_data_exfiltration_risk_hist.png', dpi=150)
plt.close()
print("Создан график: 02_data_exfiltration_risk_hist.png")

# --- График 3: Гистограммы agent_autonomy_level и resource_sensitivity ---
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

axes[0].hist(df['agent_autonomy_level'], bins=range(1, 7), edgecolor='black', alpha=0.7, color='mediumseagreen')
axes[0].set_xlabel('Agent Autonomy Level', fontsize=12)
axes[0].set_ylabel('Количество', fontsize=12)
axes[0].set_title('Распределение Agent Autonomy Level', fontsize=14, fontweight='bold')
axes[0].set_xticks(range(1, 6))

axes[1].hist(df['resource_sensitivity'], bins=range(1, 7), edgecolor='black', alpha=0.7, color='mediumpurple')
axes[1].set_xlabel('Resource Sensitivity', fontsize=12)
axes[1].set_ylabel('Количество', fontsize=12)
axes[1].set_title('Распределение Resource Sensitivity', fontsize=14, fontweight='bold')
axes[1].set_xticks(range(1, 6))

plt.tight_layout()
plt.savefig(f'{fig_dir}/03_autonomy_sensitivity_hist.png', dpi=150)
plt.close()
print("Создан график: 03_autonomy_sensitivity_hist.png")

# --- График 4: Круговая диаграмма access_decision ---
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

decision_counts = df['access_decision'].value_counts()
colors_pie = ['#ff6b6b', '#51cf66', '#ffd43b']

axes[0].pie(decision_counts.values, labels=decision_counts.index, autopct='%1.1f%%', 
            colors=colors_pie, explode=[0.05]*len(decision_counts))
axes[0].set_title('Распределение Access Decision (круговая)', fontsize=14, fontweight='bold')

# Столбчатая диаграмма
decision_colors = {'Blocked': '#ff6b6b', 'Allowed': '#51cf66', 'Needs_Human_Approval': '#ffd43b'}
bars = axes[1].bar(decision_counts.index, decision_counts.values, 
                   color=[decision_colors.get(x, 'gray') for x in decision_counts.index])
axes[1].set_xlabel('Access Decision', fontsize=12)
axes[1].set_ylabel('Количество', fontsize=12)
axes[1].set_title('Распределение Access Decision (столбчатая)', fontsize=14, fontweight='bold')
for bar in bars:
    height = bar.get_height()
    axes[1].text(bar.get_x() + bar.get_width()/2., height, f'{height}', ha='center', va='bottom')

plt.tight_layout()
plt.savefig(f'{fig_dir}/04_access_decision_pie_bar.png', dpi=150)
plt.close()
print("Создан график: 04_access_decision_pie_bar.png")

# --- График 5: Топ-5 agent_role и user_role ---
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

top_agent_roles = df['agent_role'].value_counts().head(5)
top_user_roles = df['user_role'].value_counts().head(5)

bars1 = axes[0].barh(top_agent_roles.index, top_agent_roles.values, color='steelblue')
axes[0].set_xlabel('Количество запросов', fontsize=12)
axes[0].set_title('Топ-5 Agent Role', fontsize=14, fontweight='bold')
for i, v in enumerate(top_agent_roles.values):
    axes[0].text(v + max(top_agent_roles.values)*0.01, i, str(v), va='center')

bars2 = axes[1].barh(top_user_roles.index, top_user_roles.values, color='coral')
axes[1].set_xlabel('Количество запросов', fontsize=12)
axes[1].set_title('Топ-5 User Role', fontsize=14, fontweight='bold')
for i, v in enumerate(top_user_roles.values):
    axes[1].text(v + max(top_user_roles.values)*0.01, i, str(v), va='center')

plt.tight_layout()
plt.savefig(f'{fig_dir}/05_top_roles.png', dpi=150)
plt.close()
print("Создан график: 05_top_roles.png")

# --- График 6: Тепловая карта корреляций ---
fig, ax = plt.subplots(figsize=(12, 10))
numeric_for_corr = df[['action_risk_score', 'data_exfiltration_risk', 'agent_autonomy_level', 
                       'resource_sensitivity', 'permission_match', 'prompt_injection_detected',
                       'previous_failed_attempts', 'human_approval_required', 'audit_log_available']]
corr_matrix = numeric_for_corr.corr()
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0, fmt='.2f', 
            square=True, linewidths=0.5, ax=ax)
ax.set_title('Тепловая карта корреляций числовых признаков', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{fig_dir}/06_correlation_heatmap.png', dpi=150)
plt.close()
print("Создан график: 06_correlation_heatmap.png")

# --- График 7: Boxplot action_risk_score по access_decision ---
fig, ax = plt.subplots(figsize=(10, 6))
order = ['Blocked', 'Needs_Human_Approval', 'Allowed']
sns.boxplot(data=df, x='access_decision', y='action_risk_score', order=order, 
            palette={'Blocked': '#ff6b6b', 'Needs_Human_Approval': '#ffd43b', 'Allowed': '#51cf66'}, ax=ax)
ax.set_xlabel('Access Decision', fontsize=12)
ax.set_ylabel('Action Risk Score', fontsize=12)
ax.set_title('Распределение Action Risk Score по группам Access Decision', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{fig_dir}/07_boxplot_risk_by_decision.png', dpi=150)
plt.close()
print("Создан график: 07_boxplot_risk_by_decision.png")

# --- График 8: % блокировок по комбинациям agent_role + resource_sensitivity ---
fig, ax = plt.subplots(figsize=(14, 8))
blocked_pct = df.groupby(['agent_role', 'resource_sensitivity']).apply(
    lambda x: (x['access_decision'] == 'Blocked').sum() / len(x) * 100
).unstack(fill_value=0)

sns.heatmap(blocked_pct, annot=True, fmt='.1f', cmap='YlOrRd', ax=ax, cbar_kws={'label': '% Blocked'})
ax.set_xlabel('Resource Sensitivity', fontsize=12)
ax.set_ylabel('Agent Role', fontsize=12)
ax.set_title('% Блокировок по комбинациям Agent Role и Resource Sensitivity', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{fig_dir}/08_blocked_pct_heatmap.png', dpi=150)
plt.close()
print("Создан график: 08_blocked_pct_heatmap.png")

# --- График 9: Топ-5 комбинаций (agent_role + requested_action) с наибольшим % Blocked ---
fig, ax = plt.subplots(figsize=(12, 8))
role_action_blocked = df.groupby(['agent_role', 'requested_action']).apply(
    lambda x: (x['access_decision'] == 'Blocked').sum() / len(x) * 100
).sort_values(ascending=False).head(5)

# Создаем комбинированные метки
labels = [f"{idx[0]}\n{idx[1]}" for idx in role_action_blocked.index]
bars = ax.barh(range(len(labels)), role_action_blocked.values, color='#ff6b6b')
ax.set_yticks(range(len(labels)))
ax.set_yticklabels(labels)
ax.set_xlabel('% Blocked', fontsize=12)
ax.set_title('Топ-5 комбинаций (Agent Role + Requested Action) с наибольшим % Blocked', fontsize=14, fontweight='bold')
for bar, val in zip(bars, role_action_blocked.values):
    ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2, f'{val:.1f}%', va='center')
ax.invert_yaxis()
plt.tight_layout()
plt.savefig(f'{fig_dir}/09_top_blocked_combinations.png', dpi=150)
plt.close()
print("Создан график: 09_top_blocked_combinations.png")

# --- График 10: Влияние agent_autonomy_level на решение (stacked bar) ---
fig, ax = plt.subplots(figsize=(12, 7))
autonomy_decision = df.groupby(['agent_autonomy_level', 'access_decision']).size().unstack(fill_value=0)
autonomy_decision_pct = autonomy_decision.div(autonomy_decision.sum(axis=1), axis=0) * 100

autonomy_decision_pct.plot(kind='bar', stacked=True, ax=ax, 
                           color={'Blocked': '#ff6b6b', 'Needs_Human_Approval': '#ffd43b', 'Allowed': '#51cf66'})
ax.set_xlabel('Agent Autonomy Level', fontsize=12)
ax.set_ylabel('Процент (%)', fontsize=12)
ax.set_title('Влияние Agent Autonomy Level на Access Decision', fontsize=14, fontweight='bold')
ax.legend(title='Access Decision', loc='upper right')
ax.set_xticklabels([str(int(x)) for x in autonomy_decision_pct.index], rotation=0)
plt.tight_layout()
plt.savefig(f'{fig_dir}/10_autonomy_stacked_bar.png', dpi=150)
plt.close()
print("Создан график: 10_autonomy_stacked_bar.png")

# --- График 11: Resource Type блокировки при высокой чувствительности ---
fig, ax = plt.subplots(figsize=(12, 7))
high_sens_blocked = df[(df['resource_sensitivity'] >= 4) & (df['access_decision'] == 'Blocked')]
resource_type_blocked = high_sens_blocked['resource_type'].value_counts().head(10)

bars = ax.bar(resource_type_blocked.index, resource_type_blocked.values, color='#ff6b6b', edgecolor='black')
ax.set_xlabel('Resource Type', fontsize=12)
ax.set_ylabel('Количество блокировок', fontsize=12)
ax.set_title('Топ Resource Type блокировок при высокой чувствительности (4-5)', fontsize=14, fontweight='bold')
ax.tick_params(axis='x', rotation=45)
for bar in bars:
    height = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., height, f'{int(height)}', ha='center', va='bottom')
plt.tight_layout()
plt.savefig(f'{fig_dir}/11_resource_type_blocked_high_sens.png', dpi=150)
plt.close()
print("Создан график: 11_resource_type_blocked_high_sens.png")

# --- График 12: Prompt Injection решения ---
fig, ax = plt.subplots(figsize=(10, 6))
injection_decision = injection_cases['access_decision'].value_counts()
colors_inj = {'Blocked': '#ff6b6b', 'Needs_Human_Approval': '#ffd43b', 'Allowed': '#51cf66'}
bars = ax.bar(injection_decision.index, injection_decision.values, 
              color=[colors_inj.get(x, 'gray') for x in injection_decision.index])
ax.set_xlabel('Access Decision', fontsize=12)
ax.set_ylabel('Количество', fontsize=12)
ax.set_title('Решения при обнаружении Prompt Injection', fontsize=14, fontweight='bold')
for bar in bars:
    height = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., height, f'{int(height)}', ha='center', va='bottom')
plt.tight_layout()
plt.savefig(f'{fig_dir}/12_prompt_injection_decisions.png', dpi=150)
plt.close()
print("Создан график: 12_prompt_injection_decisions.png")

print("\n" + "="*80)
print("Все визуализации созданы в папке:", fig_dir)
print("="*80)

# ============================================
# 3. ОПИСАНИЕ ИНСАЙТОВ
# ============================================
print("\n" + "="*80)
print("ФОРМИРОВАНИЕ ОТЧЕТА С ИНСАЙТАМИ")
print("="*80)

# Создаем текстовый отчет
report_lines = []
report_lines.append("="*80)
report_lines.append("ОТЧЕТ ПО EDA АНАЛИЗУ СИСТЕМЫ БЕЗОПАСНОСТИ AI-АГЕНТОВ")
report_lines.append("="*80)
report_lines.append("")

# Отчет по каждому графику
report_lines.append("""
ГРАФИК 1: Гистограмма action_risk_score
Тип: Гистограмма с пороговыми линиями
Что показывает: Распределение оценок риска действий с отметкой порогов 50 и 80
Инсайт для бизнеса: Позволяет оценить долю высокорисковых операций и настроить пороги блокировок
""")

report_lines.append("""
ГРАФИК 2: Гистограмма data_exfiltration_risk  
Тип: Гистограмма
Что показывает: Распределение риска эксфильтрации данных
Инсайт для бизнеса: Выявление паттернов потенциальной утечки данных для усиления контроля
""")

report_lines.append("""
ГРАФИК 3: Гистограммы agent_autonomy_level и resource_sensitivity
Тип: Две гистограммы
Что показывает: Распределение уровней автономии агентов и чувствительности ресурсов
Инсайт для бизнеса: Понимание баланса между автономностью агентов и критичностью ресурсов
""")

report_lines.append("""
ГРАФИК 4: Распределение access_decision
Тип: Круговая + столбчатая диаграмма
Что показывает: Соотношение Blocked/Allowed/Needs_Human_Approval решений
Инсайт для бизнеса: Оценка эффективности текущей политики доступа и нагрузки на human approval
""")

report_lines.append("""
ГРАФИК 5: Топ-5 agent_role и user_role
Тип: Горизонтальные столбчатые диаграммы
Что показывает: Наиболее активные роли агентов и пользователей
Инсайт для бизнеса: Фокус на мониторинге наиболее активных ролей для оптимизации безопасности
""")

report_lines.append("""
ГРАФИК 6: Тепловая карта корреляций
Тип: Heatmap корреляций
Что показывает: Взаимосвязи между числовыми признаками
Инсайт для бизнеса: Выявление скрытых зависимостей для улучшения модели оценки риска
""")

report_lines.append("""
ГРАФИК 7: Boxplot action_risk_score по access_decision
Тип: Boxplot
Что показывает: Распределение risk score по типам решений
Инсайт для бизнеса: Валидация согласованности решений системы с оценками риска
""")

report_lines.append("""
ГРАФИК 8: % блокировок по agent_role + resource_sensitivity
Тип: Тепловая карта
Что показывает: Процент блокировок для комбинаций ролей агентов и чувствительности ресурсов
Инсайт для бизнеса: Идентификация高风险 комбинаций для точечной настройки политик
""")

report_lines.append("""
ГРАФИК 9: Топ-5 комбинаций с наибольшим % Blocked
Тип: Горизонтальная столбчатая диаграмма
Что показывает: Комбинации агент-действие с максимальным процентом блокировок
Инсайт для бизнеса: Приоритетные области для аудита и пересмотра политик доступа
""")

report_lines.append("""
ГРАФИК 10: Влияние agent_autonomy_level на решение
Тип: Stacked bar chart
Что показывает: Как уровень автономии влияет на распределение решений
Инсайт для бизнеса: Оценка необходимости ограничения автономности для определенных сценариев
""")

report_lines.append("""
ГРАФИК 11: Resource Type блокировки при высокой чувствительности
Тип: Столбчатая диаграмма
Что показывает: Какие типы ресурсов чаще блокируются при высокой чувствительности
Инсайт для бизнеса: Фокус защиты наиболее критичных ресурсов
""")

report_lines.append("""
ГРАФИК 12: Prompt Injection решения
Тип: Столбчатая диаграмма
Что показывает: Распределение решений при обнаружении инъекций промптов
Инсайт для бизнеса: Оценка эффективности детектирования и возможных ложных срабатываний
""")

# Сохраняем отчет
with open('/workspace/EDA_Insights_Report.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(report_lines))

print("\nОтчет сохранен: /workspace/EDA_Insights_Report.txt")

# ============================================
# ВЫВОД КЛЮЧЕВЫХ СТАТИСТИК
# ============================================
print("\n" + "="*80)
print("КЛЮЧЕВЫЕ СТАТИСТИКИ")
print("="*80)

print(f"\nВсего записей: {len(df)}")
print(f"\nРаспределение access_decision:")
print(df['access_decision'].value_counts())
print(f"\nПроцент блокировок: {(df['access_decision'] == 'Blocked').mean()*100:.2f}%")
print(f"Процент разрешений: {(df['access_decision'] == 'Allowed').mean()*100:.2f}%")
print(f"Процент human approval: {(df['access_decision'] == 'Needs_Human_Approval').mean()*100:.2f}%")

print(f"\nАномалии (action_risk_score > 80 и Allowed): {len(high_risk_allowed)} случаев")
print(f"Prompt injection detected: {len(injection_cases)} случаев ({len(injection_cases)/len(df)*100:.2f}%)")
print(f"Выбросы data_exfiltration_risk (95-й перцентиль): {len(exfil_outliers)} случаев")
print(f"High failed attempts при Allowed: {len(high_attempts_success)} случаев")
print(f"Permission mismatches: {len(mismatch_df)} случаев")

print("\n" + "="*80)
print("АНАЛИЗ ЗАВЕРШЕН")
print("="*80)
print(f"\nСозданные файлы:")
print(f"  1. Excel: {excel_file}")
print(f"  2. Визуализации: {fig_dir}/")
print(f"  3. Отчет: /workspace/EDA_Insights_Report.txt")
